import requests
#from background import keep_alive
from time import sleep
import json
import telebot
import os

bot = telebot.TeleBot(token='6226283214:AAG__22ZPjTDqJRy61ABFRj6XFX67sxcoVw')


def get_json():
    i = 0
    while True:
        i += 1
        headers = {
            'authority': 'www.olx.ua',
            'accept': '*/*',
            'accept-language': 'uk',
            'authorization': 'Bearer ccbeda69abccd020feeb186dc993d98667a8613f',
            'referer':
            'https://www.olx.ua/d/uk/elektronika/kompyutery-i-komplektuyuschie/komplektuyuschie-i-aksesuary/?currency=UAH&search%5Border%5D=created_at:desc&search%5Bfilter_enum_subcategory%5D%5B0%5D=videokarty&view=list',
            'sec-ch-ua': '"Not?A_Brand";v="99", "Opera";v="97", "Chromium";v="111"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'sec-fetch-dest': 'empty',
            'sec-fetch-mode': 'cors',
            'sec-fetch-site': 'same-origin',
            'user-agent':
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/111.0.0.0 Safari/537.36 OPR/97.0.0.0',
            'x-client': 'DESKTOP',
            'x-device-id': '689af193-441a-4dfc-9cb0-ecf432e7edc0',
            'x-fingerprint':
            'fbdc4f53959cdb4ade2472495277974c5c6324e638a283d8d11dae88fc3236ac3fef60c9cf99daee3fef60c9cf99daeefaed307981c3a6ca4e1f7a2acddfea3321f2c817b2cc220c3fef60c9cf99daee3fef60c9cf99daee4e1f7a2acddfea33b497a357830277b800d2d70001a8f455308e012c59cf7bdd93ba89d2bc096e295c6324e638a283d85a1778be62509b00e8380ebb100f22a65838a5ba78a32a1b00e97cab1dc631e2f6e9813aa2a88b6b24488822f5af18b0aa6925b0558d723aa2dab357225d0511272ec8e83d084c3e6255da10575393646255da105753936400ab77cc9433c49756b16d11aecc8186428983c35d9b74ca131d7dd9eb490fb0fb7dd9ca7e63276c82beb9cbe2697dfd7fce0dd7835f6d117bf9fa6fdad87a6df0e600f8bcbb93ea00dd0ba2190d70f74348c250bb7bd9817cf587ca35748f297cf587ca35748f297cf587ca35748f297cf587ca35748f297cf587ca35748f297cf587ca35748f297cf587ca35748f297cf587ca35748f297cf587ca35748f297cf587ca35748f29e4788261c6c83237',
            'x-platform-type': 'mobile-html5',
        }
        count = i * 40
        params = {
            'offset': f'{count}',
            'limit': '40',
            'category_id': '78',
            'currency': 'UAH',
            'filter_float_price:to': '5501',
            'filter_refiners': 'spell_checker',
            'facets': '[{"field":"region","fetchLabel":true,"fetchUrl":true,"limit":30}]',
        }

        response = requests.get('https://www.olx.ua/api/v1/offers/',
                                params=params,
                                headers=headers)
        data = response.json()
        get_offers(data)
        self_link = data["links"]["self"]["href"]
        print(self_link)
        sleep(5)
        if i % 25 == 0:
            i = 0


def get_offer(item):
    # pass
    offer = {}

    offer["url"] = item["url"]
    offer["ID_advert"] = item["id"]
    offer["title"] = item["title"]
    offer["description"] = item["description"].replace('<br />', '')
    if len(offer["description"]) > 3000:
        offer["description"] = offer["description"][:3000]
    offer["price"] = item["params"][0]["value"]["value"]
    if isinstance(item["params"][0]["value"]["previous_value"], int):
        offer["previous_price"] = item["params"][0]["value"]["previous_value"]
    else:
        offer["previous_price"] = 0
    offer["last_refresh_time"] = item["last_refresh_time"]
    offer["created_time"] = item["created_time"]
    offer["city"] = item["location"]["city"]["name"]
    offer["region"] = item["location"]["region"]["name"]

    return offer


def check_database(item):
    offer = get_offer(item)

    file_name = "database.json"

    with open(file_name, 'r') as file:
        database = json.load(file)

    # database.append(offer)
    track = None
    for i in range(len(database)):
        if offer["ID_advert"] == database[i]["ID_advert"] and offer["price"] == database[i]["price"]:
            track = True
            # print('Такое есть')
            break
        elif offer["ID_advert"] == database[i]["ID_advert"] and offer["price"] != database[i]["price"]:
            database[i] = offer
            track = False
            print('Не та цена')
            break
    if track == False:
        # telegram_message(offer)
        with open('database.json', 'w') as file:
            json.dump(database, file, indent=4)
    elif track == None:
        print('Новинка')
        # telegram_message(offer)
        database.append(offer)
        with open('database.json', 'w') as file:
            json.dump(database, file, indent=4)






def telegram_message(offer):
    sleep(3)
    print('Отправил')
    chat_id = '-1002118087052'
    message_text = f'О ЧЕТА НОВОЕ : \n Заголовок : \n {offer["title"]} \n Описание : \n {offer["description"]} \n Цена : \n {offer["price"]} \n Предыдущая цена \n {offer["previous_price"]} \n Последнее время изменения \n {offer["last_refresh_time"]} \n Время создания : \n {offer["created_time"]} \n Город : \n {offer["city"]} \n Область : \n {offer["region"]} \n Ссылка : \n{offer["url"]}'
    bot.send_message(chat_id=chat_id, text=message_text)


def get_offers(data):
    offers = []
    for item in data["data"]:
        check_database(item)
        # break

    return offers


#keep_alive()


def main():
    get_json()


if __name__ == '__main__':
    filename = "database.json"
    if os.path.isfile(filename):
        print('YES YES')
        main()
    else:
        database = []
        with open("database.json", "w") as json_file:
            json.dump(database, json_file)
        main()
