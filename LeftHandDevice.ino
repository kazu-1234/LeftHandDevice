/*
 * LeftHandDevice.ino
 * v1.3.0
 * 
 * Raspberry Pi Pico 2 W 用 左手デバイスファームウェア
 * ボタンとLEDを用いたキーボード入力エミュレーション
 */

#include <Keyboard.h>
#include <EEPROM.h>
#include <MouseAbsolute.h>

// --- 定数定義 ---
#define EEPROM_SIZE 4096
#define CONFIG_MAGIC 0x1A2B3E02 // 設定データの有効性チェック用マジックナンバー (マクロ対応のため更新)

// --- ピン定義 ---
const int PIN_BTN_MODE = 10;
const int PIN_LED_MODE = 21;

// ボタン1〜5のピン
const int PIN_BTN[5] = {11, 12, 13, 14, 15};
const int PIN_BTN_LED[5] = {20, 19, 18, 17, 16};

// LEDウェーブ用配列 (モードLEDを含めた全6個)
const int ALL_LEDS[6] = {21, 20, 19, 18, 17, 16};

// --- グローバル変数 ---
bool isModeB = false; // 現在の選択モード (false: モードA単発, true: モードB連続)

// =============================================
// アサイン設定用データ構造
// =============================================
enum ActionType {
  ACTION_NONE = 0,
  ACTION_KEY = 1,
  ACTION_MOUSE = 2,
  ACTION_CMD = 3,       // アプリ起動(Win+R 経由のマクロ)
  ACTION_WAIT = 4       // 待機アクション
};

struct MacroStep {
  uint8_t type;         // ActionType
  uint8_t modifiers;    // 修飾キーのビットフラグ
  char key;             // メインキー (ASCII)
  int16_t mouseX;       // マウス絶対位置 X
  int16_t mouseY;       // マウス絶対位置 Y
  uint16_t waitMs;      // 待機時間(ms)
  char cmdString[64];   // アプリ起動用コマンド文字列 (最大63文字+Null)
};

struct ButtonMacro {
  uint8_t stepCount;    // 有効なステップ数 (0-10)
  MacroStep steps[10];  // 最大10ステップ
};

struct DeviceConfig {
  uint32_t magic;
  ButtonMacro buttons[5];
};

DeviceConfig config;

// =============================================
// ボタン状態管理用構造体
// =============================================// ボタンの状態管理構造体
struct ButtonState {
  bool prevPhysState;
  bool isRepeating;
  bool ledShouldBeOn;
  unsigned long lastPressTime;
  unsigned long lastRepeatTime;
  
  // 保存時のLED通知用フラグとタイマー
  bool isFlashing;
  uint8_t flashCount;
  unsigned long lastFlashTime;
};
ButtonState btnStates[5];
bool prevModeBtnPressed = false;

// --- 設定値 ---
const unsigned long DEBOUNCE_DELAY = 50;
const unsigned long MODEA_LED_TIME = 300;
const unsigned long REPEAT_DELAY = 50;

// =============================================
// ユーティリティ関数
// =============================================

// EEPROMから設定を読み込む
void loadConfig() {
  EEPROM.get(0, config);
  if (config.magic != CONFIG_MAGIC) {
    // 初期化 (デフォルト設定)
    config.magic = CONFIG_MAGIC;
    char defaultKeys[5] = {'a', 'b', 'c', 'd', 'e'};
    for (int i = 0; i < 5; i++) {
      config.buttons[i].stepCount = 1; // デフォルトでは1ステップのみ
      // 全ステップをゼロ初期化
      for (int s = 0; s < 10; s++) {
        config.buttons[i].steps[s].type = ACTION_NONE;
        config.buttons[i].steps[s].modifiers = 0;
        config.buttons[i].steps[s].key = 0;
        config.buttons[i].steps[s].mouseX = 0;
        config.buttons[i].steps[s].mouseY = 0;
        config.buttons[i].steps[s].waitMs = 0;
        memset(config.buttons[i].steps[s].cmdString, 0, sizeof(config.buttons[i].steps[s].cmdString));
      }
      // Step 1 だけキー設定を入れる
      config.buttons[i].steps[0].type = ACTION_KEY;
      config.buttons[i].steps[0].key = defaultKeys[i];
    }
    EEPROM.put(0, config);
    EEPROM.commit();
  }
}

// EEPROMへ設定を保存
void saveConfig() {
  EEPROM.put(0, config);
  EEPROM.commit();
}

// キー送信処理 (修飾キー対応とシーケンス実行)
void sendKeyAction(int btnIndex) {
  ButtonMacro* macro = &config.buttons[btnIndex];
  
  for (int i = 0; i < macro->stepCount; i++) {
    MacroStep* step = &macro->steps[i];
    
    if (step->type == ACTION_KEY) {
      uint8_t mods = step->modifiers;
      if (mods & 1) Keyboard.press(KEY_LEFT_CTRL);
      if (mods & 2) Keyboard.press(KEY_LEFT_SHIFT);
      if (mods & 4) Keyboard.press(KEY_LEFT_ALT);
      
      if (step->key != 0) {
        Keyboard.press(step->key);
      }
      
      delay(10); // 短時間だけ押下状態を維持
      Keyboard.releaseAll();
      
    } else if (step->type == ACTION_MOUSE) {
      // マウス絶対座標のクリック
      MouseAbsolute.move(step->mouseX, step->mouseY);
      delay(20);
      MouseAbsolute.click(MOUSE_LEFT);
      
    } else if (step->type == ACTION_WAIT) {
      // 待機アクション
      delay(step->waitMs);
      
    } else if (step->type == ACTION_CMD) {
      // アプリ起動(Win+Rマクロ)
      Keyboard.press(KEY_LEFT_GUI);
      Keyboard.press('r');
      delay(50);
      Keyboard.releaseAll();
      delay(200); // Runダイアログが立ち上がるのを少し待つ
      
      Keyboard.print(step->cmdString);
      delay(50);
      
      Keyboard.press(KEY_RETURN);
      delay(50);
      Keyboard.releaseAll();
    }
    
    // ステップ進行の間に最低限の隙間を確保
    delay(10);
  }
}

// =============================================
// Setup
// =============================================
void setup() {
  Serial.begin(115200);
  Keyboard.begin();
  MouseAbsolute.begin();
  EEPROM.begin(EEPROM_SIZE);
  loadConfig();

  // ピンの初期化
  pinMode(PIN_BTN_MODE, INPUT_PULLUP);
  pinMode(PIN_LED_MODE, OUTPUT);
  digitalWrite(PIN_LED_MODE, LOW);

  unsigned long currentTime = millis();

  for (int i = 0; i < 5; i++) {
    pinMode(PIN_BTN[i], INPUT_PULLUP);
    pinMode(PIN_BTN_LED[i], OUTPUT);
    digitalWrite(PIN_BTN_LED[i], LOW);

    btnStates[i].prevPhysState = false;
    btnStates[i].isRepeating = false;
    btnStates[i].ledShouldBeOn = false;
    btnStates[i].lastPressTime = currentTime;
    btnStates[i].lastRepeatTime = currentTime;
    btnStates[i].isFlashing = false;
    btnStates[i].flashCount = 0;
    btnStates[i].lastFlashTime = 0;
  }

  // 起動時のウェーブアニメーション (約1秒で完結)
  int waveDelay = 90;

  // 往路
  for (int i = 0; i < 6; i++) {
    digitalWrite(ALL_LEDS[i], HIGH);
    delay(waveDelay);
    digitalWrite(ALL_LEDS[i], LOW);
  }
  // 復路
  for (int i = 4; i >= 0; i--) {
    digitalWrite(ALL_LEDS[i], HIGH);
    delay(waveDelay);
    digitalWrite(ALL_LEDS[i], LOW);
  }
}

void loop() {
  unsigned long currentTime = millis();

  // =============================================
  // 1. モード切替ボタンの処理 (GP10)
  // =============================================
  bool currentModeBtn = (digitalRead(PIN_BTN_MODE) == LOW);
  if (currentModeBtn && !prevModeBtnPressed) {
    isModeB = !isModeB;
    delay(DEBOUNCE_DELAY);
  }
  prevModeBtnPressed = currentModeBtn;

  // =============================================
  // 2. 各ボタンのロジック処理（LED出力はここでは行わない）
  // =============================================
  for (int i = 0; i < 5; i++) {
    bool currentPhysState = (digitalRead(PIN_BTN[i]) == LOW);

    // ボタンが新しく押された判定
    if (currentPhysState &&
        !btnStates[i].prevPhysState &&
        (currentTime - btnStates[i].lastPressTime > DEBOUNCE_DELAY)) {

      btnStates[i].lastPressTime = currentTime;

      // リピート中 → もう一度押したら必ず停止
      if (btnStates[i].isRepeating) {
        btnStates[i].isRepeating = false;
        btnStates[i].ledShouldBeOn = false;
      } else {
        // 停止中 → 切替ボタンの現在のモードを書き込む
        if (isModeB) {
          // モードB（連続）として開始
          btnStates[i].isRepeating = true;
          sendKeyAction(i);
          btnStates[i].lastRepeatTime = currentTime;
          btnStates[i].ledShouldBeOn = true;
        } else {
          // モードA（単発）として実行
          sendKeyAction(i);
          btnStates[i].ledShouldBeOn = true;
        }
      }
    }

    // タイマーによる継続処理
    if (btnStates[i].isRepeating) {
      // モードB：リピート中の連打処理
      if (currentTime - btnStates[i].lastRepeatTime >= REPEAT_DELAY) {
        sendKeyAction(i);
        btnStates[i].lastRepeatTime = currentTime;
      }
    } else if (btnStates[i].ledShouldBeOn && !btnStates[i].isFlashing) {
      // モードAの消灯タイマー
      if (currentTime - btnStates[i].lastPressTime >= MODEA_LED_TIME) {
        btnStates[i].ledShouldBeOn = false;
      }
    }
    
    // 保存通知用LED点滅処理 (2回点滅 = ON->OFF->ON->OFF = 4 phases)
    // 1秒間で2回光らせるため、1000ms / 4 = 250ms 間隔とする
    if (btnStates[i].isFlashing) {
      if (currentTime - btnStates[i].lastFlashTime >= 250) { // 250ms間隔
        btnStates[i].lastFlashTime = currentTime;
        btnStates[i].flashCount++;

        if (btnStates[i].flashCount >= 4) {
          // 点滅終了
          btnStates[i].isFlashing = false;
          btnStates[i].flashCount = 0;
          btnStates[i].ledShouldBeOn = false;
        } else {
          // 点滅状態をトグル
          btnStates[i].ledShouldBeOn = (btnStates[i].flashCount % 2 == 0);
        }
      }
    }

    // 物理状態を更新
    btnStates[i].prevPhysState = currentPhysState;
  }

  // =============================================
  // 3. PCからのシリアル通信処理 (設定の受信と保存)
  // =============================================
  if (Serial.available() > 0) {
    String cmd = Serial.readStringUntil('\n');
    cmd.trim();
    
    // 全消去コマンド: CLEAR:ボタン番号
    if (cmd.startsWith("CLEAR:")) {
      int btnIndex = cmd.substring(6).toInt() - 1;
      if (btnIndex >= 0 && btnIndex < 5) {
        config.buttons[btnIndex].stepCount = 0;
        for (int s = 0; s < 10; s++) {
          config.buttons[btnIndex].steps[s].type = ACTION_NONE;
        }
        saveConfig();
        Serial.println("OK:CLEARED_" + String(btnIndex + 1));
      }
    }
    
    // ステップ数設定コマンド: SETCOUNT:ボタン番号:ステップ数
    else if (cmd.startsWith("SETCOUNT:")) {
      int c1 = cmd.indexOf(':');
      int c2 = cmd.indexOf(':', c1 + 1);
      if (c1 != -1 && c2 != -1) {
        int btnIndex = cmd.substring(c1 + 1, c2).toInt() - 1;
        int count = cmd.substring(c2 + 1).toInt();
        if (btnIndex >= 0 && btnIndex < 5 && count >= 0 && count <= 10) {
          config.buttons[btnIndex].stepCount = count;
          saveConfig();
          Serial.println("OK:COUNT_SAVED_" + String(btnIndex + 1));
        }
      }
    }
    
    // ステップ設定コマンド: SET:ボタン番号(1〜5):ステップインデックス(0〜9):タイプ(KEY/MOUSE/CMD/WAIT):データ
    else if (cmd.startsWith("SET:")) {
      int c1 = cmd.indexOf(':');
      int c2 = cmd.indexOf(':', c1 + 1);
      int c3 = cmd.indexOf(':', c2 + 1);
      int c4 = cmd.indexOf(':', c3 + 1);
      
      if (c1 != -1 && c2 != -1 && c3 != -1 && c4 != -1) {
        int btnIndex = cmd.substring(c1 + 1, c2).toInt() - 1;
        int stepIndex = cmd.substring(c2 + 1, c3).toInt(); // 0-based
        String typeStr = cmd.substring(c3 + 1, c4);
        String dataStr = cmd.substring(c4 + 1);
        
        if (btnIndex >= 0 && btnIndex < 5 && stepIndex >= 0 && stepIndex < 10) {
          MacroStep* step = &config.buttons[btnIndex].steps[stepIndex];
          
          if (typeStr == "KEY") {
            step->type = ACTION_KEY;
            step->modifiers = 0;
            step->key = 0;
            
            // 修飾キーの解析
            if (dataStr.indexOf("Ctrl+") != -1) { step->modifiers |= 1; dataStr.replace("Ctrl+", ""); }
            if (dataStr.indexOf("Shift+") != -1) { step->modifiers |= 2; dataStr.replace("Shift+", ""); }
            if (dataStr.indexOf("Alt+") != -1) { step->modifiers |= 4; dataStr.replace("Alt+", ""); }
            
            // メインキー文字の解析
            if (dataStr.length() > 0) {
              if (dataStr == "Enter") {
                step->key = KEY_RETURN;
              } else if (dataStr == "Esc") {
                step->key = KEY_ESC;
              } else if (dataStr == "Space") {
                step->key = ' ';
              } else {
                step->key = dataStr.charAt(0);
                if (step->key >= 'A' && step->key <= 'Z') {
                  step->key += 32;
                }
              }
            }
            saveConfig();
            Serial.println("OK:KEY_SAVED_" + String(btnIndex + 1) + "_" + String(stepIndex));
            
          } else if (typeStr == "MOUSE") {
            step->type = ACTION_MOUSE;
            int commaIdx = dataStr.indexOf(',');
            if (commaIdx != -1) {
              step->mouseX = dataStr.substring(0, commaIdx).toInt();
              step->mouseY = dataStr.substring(commaIdx + 1).toInt();
              saveConfig();
              Serial.println("OK:MOUSE_SAVED_" + String(btnIndex + 1) + "_" + String(stepIndex));
            } else {
              Serial.println("ERR:INVALID_MOUSE_FORMAT");
            }
            
          } else if (typeStr == "CMD") {
            step->type = ACTION_CMD;
            memset(step->cmdString, 0, sizeof(step->cmdString));
            // コピー長さを制限(最大63文字)
            int cpyLen = dataStr.length();
            if (cpyLen > 63) cpyLen = 63;
            dataStr.toCharArray(step->cmdString, cpyLen + 1);
            saveConfig();
            Serial.println("OK:CMD_SAVED_" + String(btnIndex + 1) + "_" + String(stepIndex));
            
          } else if (typeStr == "WAIT") {
            step->type = ACTION_WAIT;
            step->waitMs = dataStr.toInt();
            saveConfig();
            Serial.println("OK:WAIT_SAVED_" + String(btnIndex + 1) + "_" + String(stepIndex));
          }
        }
      }
    }
    
    // 意図的な通知LED点滅コマンド: FLASH:ボタン番号
    else if (cmd.startsWith("FLASH:")) {
      int btnIndex = cmd.substring(6).toInt() - 1;
      if (btnIndex >= 0 && btnIndex < 5) {
        btnStates[btnIndex].isFlashing = true;
        btnStates[btnIndex].flashCount = 0;
        btnStates[btnIndex].lastFlashTime = millis();
        btnStates[btnIndex].ledShouldBeOn = true;
        Serial.println("OK:FLASH_" + String(btnIndex + 1));
      }
    }
  }

  // =============================================
  // 4. 全LEDの状態を一括出力
  // =============================================
  digitalWrite(PIN_LED_MODE, isModeB ? HIGH : LOW);
  for (int i = 0; i < 5; i++) {
    digitalWrite(PIN_BTN_LED[i], btnStates[i].ledShouldBeOn ? HIGH : LOW);
  }

  delay(1);
}
