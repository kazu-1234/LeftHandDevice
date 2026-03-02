/*
 * LeftHandDevice.ino
 * v1.14.0
 * 
 * Raspberry Pi Pico 2 W 用 左手デバイスファームウェア
 * 同時押し・複数回押し・EEPROM可変パターン対応版
 */

#include <Keyboard.h>
#include <EEPROM.h>
#include <MouseAbsolute.h>

#ifndef KEY_KP_1
#define KEY_KP_1 0xE1
#define KEY_KP_2 0xE2
#define KEY_KP_3 0xE3
#define KEY_KP_4 0xE4
#define KEY_KP_5 0xE5
#define KEY_KP_6 0xE6
#define KEY_KP_7 0xE7
#define KEY_KP_8 0xE8
#define KEY_KP_9 0xE9
#define KEY_KP_0 0xEA
#endif

#define EEPROM_SIZE 4096
#define CONFIG_MAGIC 0x1A2B3E03

const int PIN_BTN_MODE = 10;
const int PIN_LED_MODE = 21;
const int PIN_BTN[5] = {11, 12, 13, 14, 15};
const int PIN_BTN_LED[5] = {20, 19, 18, 17, 16};
const int ALL_LEDS[6] = {21, 20, 19, 18, 17, 16};

bool isModeB = false;

enum ActionType {
  ACTION_NONE = 0,
  ACTION_KEY = 1,
  ACTION_MOUSE = 2,
  ACTION_CMD = 3,
  ACTION_WAIT = 4
};

struct MacroStep {
  uint8_t type;
  union {
    struct {
      uint8_t modifiers;
      char key;
    } keyData;
    struct {
      int16_t x;
      int16_t y;
    } mouseData;
    struct {
      uint16_t ms;
    } waitData;
    char cmdString[34];
  } payload;
};

struct PatternConfig {
  uint8_t triggerType; // 0=Single, 1=Sim, 2=Multi
  uint8_t param1;      // Btn 1..5
  uint8_t param2;      // Btn 1..5 or Tap Count 2..3
  uint8_t stepCount;
  uint16_t repeatInterval;
  MacroStep steps[10];
};

#define MAX_PATTERNS 10
struct DeviceConfig {
  uint32_t magic;
  uint8_t patternCount;
  PatternConfig patterns[MAX_PATTERNS];
};

DeviceConfig config;

// =============================================
// ボタン状態管理
// =============================================
struct PhysicalButtonState {
  bool isPressed;
  unsigned long lastToggleTime;
  unsigned long lastActionTime;  // モードB(連続モード)の連打用
  int tapCount;
  bool isDownReported;           // すでに別トリガーで使用されたか
  bool isActiveSimul;            // 現在同時押しトリガーの構成要素か
  bool isContinuousActive;      // このボタンが連続動作中か（独立管理）
  int continuousPatternIdx;     // 連続動作中のパターン番号
};
PhysicalButtonState pBtns[5];

// モード切替ボタン用デバウンス変数（Thunder_test方式）
uint8_t modeBtnState = HIGH;
uint8_t lastModeBtnReading = HIGH;
unsigned long lastModeBtnDebounceTime = 0;
const unsigned long MODE_BTN_DEBOUNCE = 50; // 50ms デバウンス

const unsigned long DEBOUNCE_DELAY = 15;
const unsigned long MULTITAP_WINDOW = 350; // 複数回押しの受付猶予(ms)
const unsigned long SIMUL_WINDOW = 50;     // 同時押しの受付猶予(ms)

// =============================================
// LEDウェーブ関数
// =============================================
void playLedWave() {
  int waveDelay = 90;
  for (int i = 0; i < 6; i++) {
    digitalWrite(ALL_LEDS[i], HIGH);
    delay(waveDelay);
    digitalWrite(ALL_LEDS[i], LOW);
  }
  for (int i = 4; i >= 0; i--) {
    digitalWrite(ALL_LEDS[i], HIGH);
    delay(waveDelay);
    digitalWrite(ALL_LEDS[i], LOW);
  }
  if (isModeB) digitalWrite(PIN_LED_MODE, HIGH);
}

// LED点灯関数（単発モード用：0.3秒間だけ点灯）
void briefLed(int ledPin) {
  digitalWrite(ledPin, HIGH);
  delay(300);
  digitalWrite(ledPin, LOW);
}

// LED点灯関数（同時押し単発用：2つのLEDを0.3秒間だけ同時点灯）
void briefLed2(int ledPin1, int ledPin2) {
  digitalWrite(ledPin1, HIGH);
  digitalWrite(ledPin2, HIGH);
  delay(300);
  digitalWrite(ledPin1, LOW);
  digitalWrite(ledPin2, LOW);
}

// LED点滅関数（アプリからの変更通知用：2回点滅）
void flashLed(int ledPin) {
  // 1回目: ON 150ms -> OFF 100ms
  digitalWrite(ledPin, HIGH);
  delay(150);
  digitalWrite(ledPin, LOW);
  delay(100);
  // 2回目: ON 150ms -> OFF
  digitalWrite(ledPin, HIGH);
  delay(150);
  digitalWrite(ledPin, LOW);
}

void flashLed2(int ledPin1, int ledPin2) {
  // 1回目
  digitalWrite(ledPin1, HIGH);
  digitalWrite(ledPin2, HIGH);
  delay(150);
  digitalWrite(ledPin1, LOW);
  digitalWrite(ledPin2, LOW);
  delay(100);
  // 2回目
  digitalWrite(ledPin1, HIGH);
  digitalWrite(ledPin2, HIGH);
  delay(150);
  digitalWrite(ledPin1, LOW);
  digitalWrite(ledPin2, LOW);
}

// =============================================
// EEPROM
// =============================================
void loadConfig() {
  EEPROM.get(0, config);
  if (config.magic != CONFIG_MAGIC) {
    config.magic = CONFIG_MAGIC;
    config.patternCount = 0;
    saveConfig();
  }
}

void saveConfig() {
  EEPROM.put(0, config);
  EEPROM.commit();
}

// =============================================
// アクション実行処理
// =============================================
void executePattern(int patIndex) {
  if (patIndex >= config.patternCount) return;
  PatternConfig* pat = &config.patterns[patIndex];
  
  for (int i = 0; i < pat->stepCount; i++) {
    MacroStep* step = &pat->steps[i];
    
    if (step->type == ACTION_KEY) {
      uint8_t mods = step->payload.keyData.modifiers;
      if (mods & 1) Keyboard.press(KEY_LEFT_CTRL);
      if (mods & 2) Keyboard.press(KEY_LEFT_SHIFT);
      if (mods & 4) Keyboard.press(KEY_LEFT_ALT);
      
      char k = step->payload.keyData.key;
      if (k != 0) Keyboard.press(k);
      
      delay(10);
      Keyboard.releaseAll();
      
    } else if (step->type == ACTION_MOUSE) {
      MouseAbsolute.move(step->payload.mouseData.x, step->payload.mouseData.y);
      delay(20);
      MouseAbsolute.click(MOUSE_LEFT);
      delay(10);
      
    } else if (step->type == ACTION_WAIT) {
      delay(step->payload.waitData.ms);
      
    } else if (step->type == ACTION_CMD) {
      Keyboard.press(KEY_LEFT_GUI);
      Keyboard.press('r');
      delay(50);
      Keyboard.releaseAll();
      delay(200);
      Keyboard.print(step->payload.cmdString);
      delay(50);
      Keyboard.press(KEY_RETURN);
      delay(50);
      Keyboard.releaseAll();
    }
    delay(10);
  }
}

// =============================================
// シリアル通信パーサ
// =============================================
void handleSerialCommand(char* cmdLine) {
  if (strncmp(cmdLine, "WAVE", 4) == 0) {
    playLedWave();
    return;
  }
  
  if (strncmp(cmdLine, "CLEAR_ALL", 9) == 0) {
    config.patternCount = 0;
    return;
  }
  
  if (strncmp(cmdLine, "ADD_PATTERN:", 12) == 0) {
    if (config.patternCount >= MAX_PATTERNS) return;
    char* token = strtok(cmdLine + 12, ":");
    if (!token) return;
    
    int pIdx = config.patternCount;
    config.patterns[pIdx].triggerType = atoi(token);
    config.patterns[pIdx].param1 = atoi(strtok(NULL, ":"));
    config.patterns[pIdx].param2 = atoi(strtok(NULL, ":"));
    config.patterns[pIdx].repeatInterval = atoi(strtok(NULL, ":"));
    config.patterns[pIdx].stepCount = atoi(strtok(NULL, ":"));
    config.patternCount++;
    return;
  }
  
  if (strncmp(cmdLine, "SET_STEP:", 9) == 0) {
    strtok(cmdLine, ":"); // SET_STEP
    int pIdx = atoi(strtok(NULL, ":"));
    int sIdx = atoi(strtok(NULL, ":"));
    char* typeStr = strtok(NULL, ":");
    char* dataStr = strtok(NULL, ""); 

    if (pIdx >= config.patternCount || sIdx >= 10 || dataStr == NULL) return;
    MacroStep* step = &config.patterns[pIdx].steps[sIdx];
    memset(step, 0, sizeof(MacroStep));
    
    if (strcmp(typeStr, "KEY") == 0) {
      step->type = ACTION_KEY;
      uint8_t mods = 0;
      // Modifier parsing (e.g. "Ctrl+Shift+a")
      while(char* plus = strchr(dataStr, '+')) {
        *plus = '\0';
        if(strstr(dataStr, "Ctrl")) mods |= 1;
        if(strstr(dataStr, "Shift")) mods |= 2;
        if(strstr(dataStr, "Alt")) mods |= 4;
        dataStr = plus + 1;
      }
      if (strcmp(dataStr, "Enter") == 0) step->payload.keyData.key = KEY_RETURN;
      else if (strcmp(dataStr, "Esc") == 0) step->payload.keyData.key = KEY_ESC;
      else if (strcmp(dataStr, "Space") == 0) step->payload.keyData.key = ' ';
      else if (strcmp(dataStr, "NumPad0") == 0) step->payload.keyData.key = KEY_KP_0;
      else if (strcmp(dataStr, "NumPad1") == 0) step->payload.keyData.key = KEY_KP_1;
      else if (strcmp(dataStr, "NumPad2") == 0) step->payload.keyData.key = KEY_KP_2;
      else if (strcmp(dataStr, "NumPad3") == 0) step->payload.keyData.key = KEY_KP_3;
      else if (strcmp(dataStr, "NumPad4") == 0) step->payload.keyData.key = KEY_KP_4;
      else if (strcmp(dataStr, "NumPad5") == 0) step->payload.keyData.key = KEY_KP_5;
      else if (strcmp(dataStr, "NumPad6") == 0) step->payload.keyData.key = KEY_KP_6;
      else if (strcmp(dataStr, "NumPad7") == 0) step->payload.keyData.key = KEY_KP_7;
      else if (strcmp(dataStr, "NumPad8") == 0) step->payload.keyData.key = KEY_KP_8;
      else if (strcmp(dataStr, "NumPad9") == 0) step->payload.keyData.key = KEY_KP_9;
      else if (strcmp(dataStr, "F1") == 0) step->payload.keyData.key = KEY_F1;
      else if (strcmp(dataStr, "F2") == 0) step->payload.keyData.key = KEY_F2;
      else if (strcmp(dataStr, "F3") == 0) step->payload.keyData.key = KEY_F3;
      else if (strcmp(dataStr, "F4") == 0) step->payload.keyData.key = KEY_F4;
      else if (strcmp(dataStr, "F5") == 0) step->payload.keyData.key = KEY_F5;
      else if (strcmp(dataStr, "F6") == 0) step->payload.keyData.key = KEY_F6;
      else if (strcmp(dataStr, "F7") == 0) step->payload.keyData.key = KEY_F7;
      else if (strcmp(dataStr, "F8") == 0) step->payload.keyData.key = KEY_F8;
      else if (strcmp(dataStr, "F9") == 0) step->payload.keyData.key = KEY_F9;
      else if (strcmp(dataStr, "F10") == 0) step->payload.keyData.key = KEY_F10;
      else if (strcmp(dataStr, "F11") == 0) step->payload.keyData.key = KEY_F11;
      else if (strcmp(dataStr, "F12") == 0) step->payload.keyData.key = KEY_F12;
      else step->payload.keyData.key = dataStr[0];
      step->payload.keyData.modifiers = mods;
    }
    else if (strcmp(typeStr, "MOUSE") == 0) {
      step->type = ACTION_MOUSE;
      char* comma = strchr(dataStr, ',');
      if (comma) {
        *comma = '\0';
        step->payload.mouseData.x = atoi(dataStr);
        step->payload.mouseData.y = atoi(comma + 1);
      }
    }
    else if (strcmp(typeStr, "WAIT") == 0) {
      step->type = ACTION_WAIT;
      step->payload.waitData.ms = atoi(dataStr);
    }
    else if (strcmp(typeStr, "CMD") == 0) {
      step->type = ACTION_CMD;
      strncpy(step->payload.cmdString, dataStr, 33);
    }
    return;
  }
  
  if (strncmp(cmdLine, "SAVE_CONFIG", 11) == 0) {
    saveConfig();
    return;
  }
  // WPFアプリから特定のボタンLEDを点滅させるコマンド
  // フォーマット: FLASH_BUTTONS:btn1:btn2 (btn2が-1の場合は1つのみ)
  if (strncmp(cmdLine, "FLASH_BUTTONS:", 14) == 0) {
    char cmdCopy[32];
    strncpy(cmdCopy, cmdLine, 31);
    cmdCopy[31] = '\0';
    
    strtok(cmdCopy, ":"); // "FLASH_BUTTONS"
    char* b1Str = strtok(NULL, ":");
    char* b2Str = strtok(NULL, ":");
    
    int b1 = b1Str ? atoi(b1Str) : -1;
    int b2 = b2Str ? atoi(b2Str) : -1;
    
    if (b1 >= 0 && b1 < 5 && b2 >= 0 && b2 < 5) {
      flashLed2(PIN_BTN_LED[b1], PIN_BTN_LED[b2]);
    } else if (b1 >= 0 && b1 < 5) {
      flashLed(PIN_BTN_LED[b1]);
    }
    return;
  }
}

// =============================================
// ロジック＆ループ処理
// =============================================
void setup() {
  Serial.begin(115200);
  Keyboard.begin();
  MouseAbsolute.begin();
  EEPROM.begin(EEPROM_SIZE);
  loadConfig();

  pinMode(PIN_BTN_MODE, INPUT_PULLUP);
  pinMode(PIN_LED_MODE, OUTPUT);
  digitalWrite(PIN_LED_MODE, LOW);

  unsigned long current = millis();
  for (int i = 0; i < 5; i++) {
    pinMode(PIN_BTN[i], INPUT_PULLUP);
    pinMode(PIN_BTN_LED[i], OUTPUT);
    digitalWrite(PIN_BTN_LED[i], LOW);

    pBtns[i].isPressed = false;
    pBtns[i].lastToggleTime = current;
    pBtns[i].lastActionTime = current;
    pBtns[i].tapCount = 0;
    pBtns[i].isDownReported = false;
    pBtns[i].isActiveSimul = false;
    pBtns[i].isContinuousActive = false;
    pBtns[i].continuousPatternIdx = -1;
  }

  playLedWave();
}

char serialBuffer[128];
int serialIdx = 0;

void loop() {
  unsigned long currentTime = millis();

  // シリアル受信処理
  while (Serial.available() > 0) {
    char c = Serial.read();
    if (c == '\n' || c == '\r') {
      if (serialIdx > 0) {
        serialBuffer[serialIdx] = '\0';
        handleSerialCommand(serialBuffer);
        serialIdx = 0;
      }
    } else if (serialIdx < 127) {
      serialBuffer[serialIdx++] = c;
    }
  }

  // モード切替ボタン（Thunder_test方式の安定デバウンス）
  int modeBtnReading = digitalRead(PIN_BTN_MODE);
  if (modeBtnReading != lastModeBtnReading) {
    lastModeBtnDebounceTime = currentTime;
  }
  if ((currentTime - lastModeBtnDebounceTime) > MODE_BTN_DEBOUNCE) {
    if (modeBtnReading != modeBtnState) {
      modeBtnState = modeBtnReading;
      // ボタンが押された瞬間（LOW）にモード切替
      if (modeBtnState == LOW) {
        isModeB = !isModeB;
      }
    }
  }
  lastModeBtnReading = modeBtnReading;

  // ボタン状態の更新 (デバウンス処理のみ、LEDはここでは制御しない)
  for (int i = 0; i < 5; i++) {
    bool rawState = (digitalRead(PIN_BTN[i]) == LOW);
    if (rawState != pBtns[i].isPressed && (currentTime - pBtns[i].lastToggleTime > DEBOUNCE_DELAY)) {
      pBtns[i].isPressed = rawState;
      pBtns[i].lastToggleTime = currentTime;
      
      if (rawState) {
        // 押した瞬間
        // もしこのボタンが連続動作中なら、解除する
        if (pBtns[i].isContinuousActive) {
          int stoppingPattern = pBtns[i].continuousPatternIdx;
          // このボタンを停止
          pBtns[i].isContinuousActive = false;
          pBtns[i].continuousPatternIdx = -1;
          digitalWrite(PIN_BTN_LED[i], LOW);
          pBtns[i].isDownReported = true; // この押下は解除で消費する
          
          // 同じパターンを共有する他のボタンも同時に停止（同時押しの片方解除対応）
          for (int k = 0; k < 5; k++) {
            if (k != i && pBtns[k].isContinuousActive && pBtns[k].continuousPatternIdx == stoppingPattern) {
              pBtns[k].isContinuousActive = false;
              pBtns[k].continuousPatternIdx = -1;
              digitalWrite(PIN_BTN_LED[k], LOW);
            }
          }
        } else {
          pBtns[i].tapCount++;
          pBtns[i].isDownReported = false;
          pBtns[i].isActiveSimul = false;
        }
      }
    }
  }

  // 連続動作中のボタンを処理（独立した繰り返し実行）
  for (int i = 0; i < 5; i++) {
    if (pBtns[i].isContinuousActive && pBtns[i].continuousPatternIdx >= 0) {
      int pIdx = pBtns[i].continuousPatternIdx;
      if (pIdx < config.patternCount) {
        PatternConfig* pat = &config.patterns[pIdx];
        if (pat->repeatInterval > 0 && currentTime - pBtns[i].lastActionTime >= pat->repeatInterval) {
          pBtns[i].lastActionTime = currentTime;
          executePattern(pIdx);
        }
      }
    }
  }

  // トリガー判定ロジック
  // 優先順位: 1. 同時押し, 2. 複数回押し, 3. 単発押し または 連続押し(モードB)
  for (int i = 0; i < config.patternCount; i++) {
    PatternConfig* pat = &config.patterns[i];
    int btnIdx1 = pat->param1 - 1;
    if (btnIdx1 < 0 || btnIdx1 > 4) continue;
    
    // --- 同時押し (Type 1) ---
    if (pat->triggerType == 1) {
      int btnIdx2 = pat->param2 - 1;
      if (btnIdx2 >= 0 && btnIdx2 <= 4 && btnIdx1 != btnIdx2) {
        if (pBtns[btnIdx1].isPressed && pBtns[btnIdx2].isPressed) {
          // 両方押されていて、かつどちらも消費されておらず、時間差がSIMUL_WINDOW以下
          unsigned long tDiff = (pBtns[btnIdx1].lastToggleTime > pBtns[btnIdx2].lastToggleTime) 
            ? (pBtns[btnIdx1].lastToggleTime - pBtns[btnIdx2].lastToggleTime) 
            : (pBtns[btnIdx2].lastToggleTime - pBtns[btnIdx1].lastToggleTime);
            
            if (!pBtns[btnIdx1].isDownReported && !pBtns[btnIdx2].isDownReported && tDiff <= SIMUL_WINDOW) {
            pBtns[btnIdx1].isDownReported = true;
            pBtns[btnIdx2].isDownReported = true;
            pBtns[btnIdx1].isActiveSimul = true;
            pBtns[btnIdx2].isActiveSimul = true;
            pBtns[btnIdx1].lastActionTime = currentTime;
            pBtns[btnIdx2].lastActionTime = currentTime;
            
            if (isModeB) {
              // 連続モード時：独立した連続動作を両方のボタンで開始
              executePattern(i);
              pBtns[btnIdx1].isContinuousActive = true;
              pBtns[btnIdx1].continuousPatternIdx = i;
              pBtns[btnIdx2].isContinuousActive = true;
              pBtns[btnIdx2].continuousPatternIdx = i;
              digitalWrite(PIN_BTN_LED[btnIdx1], HIGH);
              digitalWrite(PIN_BTN_LED[btnIdx2], HIGH);
            } else {
              // 単発モード時は実行後に0.3秒点灯
              executePattern(i);
              briefLed2(PIN_BTN_LED[btnIdx1], PIN_BTN_LED[btnIdx2]);
            }
          }
        }
      }
    }
    
    // --- 複数回押し (Type 2) ---
    else if (pat->triggerType == 2) {
      // 連続でTapCount回押された場合。
      int reqTaps = pat->param2; // 2 or 3
      if (!pBtns[btnIdx1].isPressed && pBtns[btnIdx1].tapCount == reqTaps) {
        // ボタンが離されていて、タップカウントが一致し、猶予時間を過ぎようとしている時に発動
        if (currentTime - pBtns[btnIdx1].lastToggleTime > SIMUL_WINDOW && 
            currentTime - pBtns[btnIdx1].lastToggleTime < MULTITAP_WINDOW && 
            !pBtns[btnIdx1].isDownReported && !pBtns[btnIdx1].isActiveSimul) {
              
          pBtns[btnIdx1].tapCount = 0; // 消費
          pBtns[btnIdx1].isDownReported = true;
          
          // 複数回押しは常に単発動作なので、0.3秒点灯
          executePattern(i);
          briefLed(PIN_BTN_LED[btnIdx1]);
        }
      }
      // 規定時間を過ぎたらタップカウントリセット
      if (pBtns[btnIdx1].tapCount > 0 && currentTime - pBtns[btnIdx1].lastToggleTime > MULTITAP_WINDOW) {
        if (!pBtns[btnIdx1].isPressed) pBtns[btnIdx1].tapCount = 0;
      }
    }
  }

  // --- 単押し (Type 0) とマルチタップのリセット処理 ---
  for (int i = 0; i < config.patternCount; i++) {
    PatternConfig* pat = &config.patterns[i];
    int btnIdx = pat->param1 - 1;
    if (btnIdx < 0 || btnIdx > 4) continue;

    if (pat->triggerType == 0) {
      if (pBtns[btnIdx].isPressed) {
        // 押された瞬間。SIMUL判定を待つため少し遅延して判定
        if (!pBtns[btnIdx].isDownReported && !pBtns[btnIdx].isActiveSimul && 
            (currentTime - pBtns[btnIdx].lastToggleTime > SIMUL_WINDOW)) {
          // 他の複数回押しのパターンが存在するか探す
          bool hasMultiTap = false;
          for(int j=0; j<config.patternCount; j++) {
            if (config.patterns[j].param1 - 1 == btnIdx && config.patterns[j].triggerType == 2) hasMultiTap = true;
          }

          bool canExecute = false;
          if (!hasMultiTap) {
            canExecute = true;
          } else {
            // マルチタップのパターンがある場合、長押しと判定
            if (currentTime - pBtns[btnIdx].lastToggleTime > MULTITAP_WINDOW) {
              canExecute = true;
            }
          }

          if (canExecute) {
            pBtns[btnIdx].isDownReported = true;
            pBtns[btnIdx].lastActionTime = currentTime;
            
            if (isModeB) {
              // 連続モード時：独立した連続動作を開始し、LED常時点灯
              executePattern(i);
              pBtns[btnIdx].isContinuousActive = true;
              pBtns[btnIdx].continuousPatternIdx = i;
              digitalWrite(PIN_BTN_LED[btnIdx], HIGH);
            } else {
              // 単発モード時：実行後に0.3秒点灯
              executePattern(i);
              briefLed(PIN_BTN_LED[btnIdx]);
            }
          }
        }
      } else {
        // 離した場合のマルチタップ判定リセット
        if (!pBtns[btnIdx].isDownReported && (currentTime - pBtns[btnIdx].lastToggleTime > MULTITAP_WINDOW)) {
           // すでに過ぎていたら何もしない
        }
      }
    }
  }

  // モード切替LEDを毎ループ末尾で確定させる（他の処理中のdelay等で不安定にならないようにする）
  digitalWrite(PIN_LED_MODE, isModeB ? HIGH : LOW);
}
