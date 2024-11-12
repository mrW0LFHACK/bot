import telebot
import pandas as pd
from collections import Counter
import matplotlib.pyplot as plt
from io import BytesIO
import sqlite3
import bot
bot = telebot.TeleBot(bot.token)

data = pd.read_excel("data.xlsx")
names = data["B"].dropna()    
reasons = data["C"].dropna()   
reviews = data["D"].fillna("") 

def get_user_settings(user_id):
    conn = sqlite3.connect('settings.db')
    cursor = conn.cursor()
    cursor.execute('SELECT chart_type FROM user_settings WHERE user_id = ?', (user_id,))
    result = cursor.fetchone()
    conn.close()
    return result[0] if result else 'pie'

def set_user_settings(user_id, chart_type):
    conn = sqlite3.connect('settings.db')
    cursor = conn.cursor()
    cursor.execute('REPLACE INTO user_settings (user_id, chart_type) VALUES (?, ?)', (user_id, chart_type))
    conn.commit()
    conn.close()

def get_top_characters(disappointing=True, top_n=10):
    name_counts = names.value_counts()
    return name_counts.head(top_n) if disappointing else name_counts.tail(top_n)

def generate_pie_chart():
    name_counts = names.value_counts()
    plt.figure(figsize=(8, 8))
    
    wedges, _ = plt.pie(name_counts, startangle=140)
    plt.title("Распределение разочарований персонажами")
    
    legend_labels = [f"{count} - {name} ({count / sum(name_counts) * 100:.1f}%)" 
                     for name, count in zip(name_counts.index, name_counts.values)]
    
    plt.legend(wedges, legend_labels, loc="upper center", bbox_to_anchor=(0.5, -0.1),
               fancybox=True, shadow=True, ncol=1, fontsize="small", title="Легенда")
    
    pie_chart = BytesIO()
    plt.savefig(pie_chart, format='png', bbox_inches='tight')
    pie_chart.seek(0)
    return pie_chart

def generate_bar_chart():
    name_counts = names.value_counts()
    plt.figure(figsize=(10, 6))
    
    bars = plt.bar(name_counts.index, name_counts.values, color='skyblue')
    plt.title("Распределение разочарований персонажами")
    plt.xlabel("Персонажи")
    plt.ylabel("Количество упоминаний")
    
    plt.xticks(rotation=90)
    bar_chart = BytesIO()
    plt.savefig(bar_chart, format='png', bbox_inches='tight')
    bar_chart.seek(0)
    return bar_chart


def get_character_feedback(character_name, max_reviews=5, max_reasons=5):
    character_data = data[data["B"] == character_name]
    reasons = character_data["C"].dropna().tolist()
    reviews = character_data["D"].dropna().tolist()
    
    reason_counts = Counter(reasons)
    most_common_reasons = reason_counts.most_common()
    most_frequent_reasons = [reason for reason, count in most_common_reasons if count == most_common_reasons[0][1]]
    
    return (reviews[:max_reviews], reasons[:max_reasons], most_frequent_reasons)

@bot.message_handler(commands=['start'])
def send_welcome(message):
    bot.reply_to(message, "Добро пожаловать! Выберите команду:\n`/топ_разочарований`\n`/топ_неразочарований`\n`/диаграмма_разочарований`\n`/отзыв_персонажа`\n`/настройки`", parse_mode="Markdown")

@bot.message_handler(commands=['топ_разочарований'])
def top_disappointing(message):
    top_characters = get_top_characters(disappointing=True)
    bot.send_message(message.chat.id, f"Топ-10 разочаровывающих персонажей:\n{top_characters}")

@bot.message_handler(commands=['топ_неразочарований'])
def top_non_disappointing(message):
    top_characters = get_top_characters(disappointing=False)
    bot.send_message(message.chat.id, f"Топ-10 неразочаровывающих персонажей:\n{top_characters}")

@bot.message_handler(commands=['диаграмма_разочарований'])
def send_chart(message):
    user_id = message.chat.id
    chart_type = get_user_settings(user_id)
    
    if chart_type == 'pie':
        chart = generate_pie_chart()
    else:
        chart = generate_bar_chart()
    
    bot.send_photo(message.chat.id, chart, caption="Общее распределение разочарований")

@bot.message_handler(commands=['отзыв_персонажа'])
def ask_for_character(message):
    bot.send_message(message.chat.id, "Введите имя персонажа, чтобы увидеть отзывы:")

@bot.message_handler(commands=['настройки'])
def settings(message):
    bot.send_message(message.chat.id, "Выберите тип диаграммы:\n`/круг` для круговой диаграммы\n`/столбцы` для столбчатой диаграммы")

@bot.message_handler(commands=['круг'])
def set_pie_chart(message):
    set_user_settings(message.chat.id, 'pie')
    bot.send_message(message.chat.id, "Теперь будут отображаться круговые диаграммы.")

@bot.message_handler(commands=['столбцы'])
def set_bar_chart(message):
    set_user_settings(message.chat.id, 'bar')
    bot.send_message(message.chat.id, "Теперь будут отображаться столбчатые диаграммы.")

@bot.message_handler(func=lambda message: True)
def send_character_feedback(message):
    try:
        character_name = message.text.strip()
        
        reviews, reasons, most_frequent_reasons = get_character_feedback(character_name)
        
        name_counts = names.value_counts()
        position = (name_counts.index == character_name).argmax() + 1 if character_name in name_counts else None
        mentions = name_counts.get(character_name, 0)
        
        reason_counts = Counter(reasons)
        reason_stats = "\n".join([f"{reason}: {count}" for reason, count in reason_counts.items()])
        
        if not reviews and not reasons:
            bot.send_message(message.chat.id, f"Отзывов для {character_name} не найдено")
        else:
            feedback_text = f"Отзывы о {character_name}:\n\n"
            
            if position:
                feedback_text += f"Позиция в рейтинге разочарования: {position}-е место ({mentions / sum(name_counts) * 100:.1f}%)\n"
            feedback_text += f"Количество упоминаний как разочаровывающий персонаж: {mentions}\n\n"
            
            if reviews:
                feedback_text += "Отзывы:\n" + "\n".join(reviews) + "\n\n"
            
            if reason_stats:
                feedback_text += f"Общая статистика причин разочарования:\n{reason_stats}\n\n"
            
            if most_frequent_reasons:
                feedback_text += f"Частые причины: {', '.join(most_frequent_reasons)}"
            
            max_length = 4096
            for i in range(0, len(feedback_text), max_length):
                bot.send_message(message.chat.id, feedback_text[i:i + max_length])
    
    except Exception as e:
        bot.send_message(message.chat.id, f"Произошла ошибка: {str(e)}")

bot.polling()
