from PIL import Image, ImageDraw

def create_icon():
    # 256x256の透明背景の画像を生成
    size = 256
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    
    # 濃いグレーの角丸背景（デバイス本体）
    margin = 20
    draw.rounded_rectangle(
        [margin, margin, size - margin, size - margin],
        radius=30,
        fill=(40, 44, 52, 255),
        outline=(100, 108, 120, 255),
        width=4
    )
    
    # 5つのボタン（サイコロのような配置または横・縦一列など）
    # 今回は十字レイアウト＋中央のような配置 (左手デバイス風)
    btn_size = 40
    btn_color = (198, 208, 245, 255) # 明るい青系（Material Blue/Amberに近い雰囲気）
    btn_outline = (130, 150, 200, 255)
    
    # サイコロの5の目（四角）のような配置
    positions = [
        (80, 80),    # 左上
        (176, 80),   # 右上
        (128, 128),  # 中央
        (80, 176),   # 左下
        (176, 176)   # 右下
    ]
    
    for px, py in positions:
        x0 = px - btn_size // 2
        y0 = py - btn_size // 2
        x1 = px + btn_size // 2
        y1 = py + btn_size // 2
        draw.rounded_rectangle(
            [x0, y0, x1, y1],
            radius=10,
            fill=btn_color,
            outline=btn_outline,
            width=3
        )

    # 保存
    img.save('app_icon.png')
    img.save('app_icon.ico', format='ICO', sizes=[(256, 256), (128, 128), (64, 64), (32, 32), (16, 16)])
    print("Icons generated: app_icon.png, app_icon.ico")

if __name__ == '__main__':
    create_icon()
