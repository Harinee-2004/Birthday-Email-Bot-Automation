import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os
import json
import win32com.client as win32
import base64
from datetime import datetime

def load_config(config_path):
    with open(config_path, 'r') as f:
        return json.load(f)

def load_employees(excel_path):
    df = pd.read_excel(excel_path)
    today = datetime.today().strftime('%m-%d')
    df['DOB_MMDD'] = pd.to_datetime(df['DOB']).dt.strftime('%m-%d')
    return df[df['DOB_MMDD'] == today]

def create_birthday_card(config, employees_today):
    bg_path = config['background_image']
    output_dir = config['output_directory']
    os.makedirs(output_dir, exist_ok=True)

    total_people = len(employees_today)
    positions = config['positions'].get(str(total_people), [])

    if not positions:
        print(f"⚠️ No placement config for {total_people} people.")
        return None

    # Load the background
    bg = Image.open(bg_path).convert("RGBA")
    composite = Image.new("RGBA", bg.size, (255, 255, 255, 0))

    for i, (_, emp) in enumerate(employees_today.iterrows()):
        pos = positions[i]
        photo_path = os.path.join(config['employee_photos_dir'], emp['Photo'])

        # Load and resize photo
        emp_img = Image.open(photo_path).convert("RGBA")
        emp_img = emp_img.resize((pos['width'], pos['height']), Image.LANCZOS)

        # Paste photo
        composite.paste(emp_img, (pos['x'], pos['y']), emp_img)

        # Draw name
        draw = ImageDraw.Draw(composite)
        try:
            font = ImageFont.truetype("arial.ttf", size=pos['font_size'])
        except:
            font = ImageFont.load_default()

        text = emp['Name']
        bbox = font.getbbox(text)
        text_width = bbox[2] - bbox[0]
        x_text = pos['x'] + (pos['width'] - text_width) / 2
        y_text = pos['y'] + pos['height'] + pos['name_offset_y']

        draw.text((x_text, y_text), text, font=font, fill=(189, 65, 84))

    final_card = Image.alpha_composite(bg, composite)
    
    original_width, original_height = final_card.size
    resized_width = int(original_width * 0.65)
    resized_height = int(original_height * 0.65)
    final_card = final_card.resize((resized_width, resized_height), Image.LANCZOS)

    output_path = os.path.join(output_dir, config['output_image'])
    final_card.save(output_path)
    print(f"✅ Birthday card saved: {output_path}")
    return output_path

def send_outlook_email(config, card_path, employees_today):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    # Dynamically get today's birthday recipients
    to_emails = [emp['Email'] for _, emp in employees_today.iterrows()]
    mail.To = "; ".join(to_emails)

    # Static BCC
    mail.BCC = "gtm-india@pega.com"

    # Subject line includes names
    subject_names = ", ".join([emp['Name'] for _, emp in employees_today.iterrows()])
    mail.Subject = f"🎉 Happy Birthday - {subject_names}!"

    with open(card_path, "rb") as img_file:
        encoded = base64.b64encode(img_file.read()).decode('utf-8')

    mail.HTMLBody = f"""
    <html>
    <body>
        <img src="data:image/png;base64,{encoded}" alt="Birthday Card" style="max-width: 100%;">
    </body>
    </html>
    """

    mail.Display(True)
    print(f"✅ Outlook draft created for: {mail.To}")


if __name__ == "__main__":
    config = load_config("NewplacementsTry.json")
    employees_today = load_employees("employees.xlsx")

    if employees_today.empty:
        print("🎉 No birthdays today!")
    else:
        card = create_birthday_card(config, employees_today)
        if card:
            send_outlook_email(config, card, employees_today)
