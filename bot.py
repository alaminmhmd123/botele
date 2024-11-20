import telebot
import requests
import pandas as pd
from io import BytesIO
from tabulate import tabulate
import os
from fuzzywuzzy import process, fuzz
import time

# أدخل هنا التوكن الخاص بالبوت الخاص بك
bot = telebot.TeleBot('758716287:AAH0qm5kLnUm4r9AaPteHQIJacx71VyUymQ')

# رابط التحميل المباشر للملف
EXCEL_URL = 'https://1drv.ms/x/c/84cafb26d1508441/EUGEUNEm-8oggITEBQAAAAABPBKjYUiBUdBKAh2dbwEYEw?download=1'

# تحميل بيانات ملف Excel وتخزينها في DataFrame
def load_excel_data(name):
    response = requests.get(EXCEL_URL)
    filexcl = BytesIO(response.content)
    df = pd.read_excel(filexcl, sheet_name=name, engine='openpyxl')
    return df

# توليد وصف الصورة بناءً على رقم الموديل
def generate_description(row):
    description = f"""
    رقم الموديل: {row['اسم الموديل']}
    الوصف: {row['الوصف']}
    الموجود: {row['الموجود']}
    الجاهز: {row['الجاهز']}
    الذي يحتاج تنزيل: {row['يحتاج تنزيل']}
    رأس المال: {row['رأس المال']}
    سعر المبيع: {row['السعر']}
    """
    return description

# وظيفة للتحقق مما إذا كانت كل الكلمات من الجملة موجودة في النص
def all_words_in_text(words, text):
    return all(word in text for word in words)

@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    bot.reply_to(message, "مرحباً بك في بوت الورشة قم بإرسال كلمة الملخص أو كلمة الكل أو صف شكل الطقم الذي تريده")

@bot.message_handler(func=lambda message: message.text.lower() == 'الملخص')
def send_summary(message):
    df = load_excel_data("ملخص")
    summary_text = tabulate(df, headers='keys', tablefmt='grid', showindex=False)
    bot.send_message(message.chat.id, f"الملخص:\n\n{summary_text}")

@bot.message_handler(func=lambda message: True)
def send_image(message):
    text = message.text.lower()
    df = load_excel_data("الصفحة الرئيسية")

    # تأكد من تحويل جميع الأوصاف إلى نصوص
    df['الوصف'] = df['الوصف'].astype(str)

    if text == "الكل":
        images = sorted(
            [f for f in os.listdir('./image/') if f.endswith('.jpg')],
            key=lambda x: int(os.path.splitext(x)[0])
        )
        for filename in images:
            model_number = os.path.splitext(filename)[0]
            row = df[df['اسم الموديل'] == int(model_number)]
            if not row.empty:
                description = generate_description(row.iloc[0])
                with open(os.path.join('./image/', filename), 'rb') as image:
                    bot.send_photo(message.chat.id, image, caption=description)
    else:
        # البحث باستخدام fuzzywuzzy وتحسين البحث عبر كلمات الجملة
        descriptions = df['الوصف'].tolist()
        search_words = text.split()  # تقسيم النص إلى كلمات
        matched_descriptions = []

        for desc in descriptions:
            if all_words_in_text(search_words, desc.lower()):
                matched_descriptions.append(desc)

        if matched_descriptions:
            # تتبع الصور التي تم إرسالها بالفعل لتجنب التكرار
            sent_images = set()

            # إرسال الصور التي تحتوي على الأوصاف المشابهة
            for matched_description in matched_descriptions:
                rows = df[df['الوصف'].str.contains(matched_description, case=False, na=False)]
                for _, row in rows.iterrows():
                    image_path = os.path.join('./image/', str(row['اسم الموديل']) + '.jpg')
                    description = generate_description(row)
                    
                    if image_path not in sent_images:
                        if os.path.exists(image_path):
                            with open(image_path, 'rb') as image:
                                bot.send_photo(message.chat.id, image, caption=description)
                            sent_images.add(image_path)
                        else:
                            bot.reply_to(message, "لم أتمكن من العثور على الصورة المطلوبة.")
        else:
            bot.reply_to(message, "لم أتمكن من العثور على وصف مشابه.")

@bot.message_handler(func=lambda message: True)
def send_default(message):
    bot.reply_to(message, "لم أتمكن من التعرف على الأمر. حاول مرة أخرى.")

# إعادة تشغيل البوت في حالة حدوث خطأ
while True:
    try:
        bot.polling()
    except Exception as e:
        print(f"حدث خطأ: {e}")
        time.sleep(5)  # انتظار 5 ثوانٍ قبل إعادة تشغيل البوت
