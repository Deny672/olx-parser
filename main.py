import requests
from time import sleep
import openpyxl
import telebot

bot = telebot.TeleBot(token='6226283214:AAG__22ZPjTDqJRy61ABFRj6XFX67sxcoVw')


def get_json():
    for i in range(1, 26):
        sleep(4)
        headers = {
            'authority': 'www.olx.ua',
            'accept': '*/*',
            'accept-language': 'uk',
            'authorization': 'Bearer ccbeda69abccd020feeb186dc993d98667a8613f',
            # 'cookie': 'mobile_default=desktop; my_city_2=284_0_0_%D0%9F%D0%BE%D0%BB%D1%82%D0%B0%D0%B2%D0%B0_0_%D0%9F%D0%BE%D0%BB%D1%82%D0%B0%D0%B2%D1%81%D1%8C%D0%BA%D0%B0+%D0%BE%D0%B1%D0%BB%D0%B0%D1%81%D1%82%D1%8C_poltava; fingerprint=MTI1NzY4MzI5MTsyOzA7MDswOzE7MDswOzA7MDswOzE7MTsxOzE7MTsxOzE7MTsxOzE7MTsxOzE7MTswOzE7MTsxOzA7MDsxOzE7MTsxOzE7MTsxOzE7MTsxOzA7MTsxOzA7MTsxOzE7MDswOzA7MDswOzA7MTswOzE7MTswOzA7MDsxOzA7MDsxOzE7MDsxOzE7MTsxOzA7MTswOzM5NjA2MDQ3NTI7MjsyOzI7MjsyOzI7NTsyODQ4MDA2NDE4OzEzNTcwNDE3Mzg7MTsxOzE7MTsxOzE7MTsxOzE7MTsxOzE7MTsxOzE7MTsxOzA7MDswOzM1NjM2MzY2NzszNDY5MzA2NTUxOzMwODQ4OTk4MzU7Nzg1MjQ3MDI5OzEwMDUzMDEyMDM7MTM2Njs3Njg7MjQ7MjQ7MTgwOzEyMDsxODA7MTIwOzE4MDsxMjA7MTgwOzEyMDsxODA7MTIwOzE4MDsxMjA7MTgwOzEyMDsxODA7MTIwOzE4MDsxMjA7MTgwOzEyMDswOzA7MA==; dfp_user_id=5127bcf3-cc94-4903-920d-6b0208880ed4-ver2; __utmc=250720985; __gads=ID=0826eefb2f250446:T=1672917975:S=ALNI_MbeUxNBeHvDczXt427bxGp2jrMU1g; cookieBarSeen=true; lister_lifecycle=1672917994; searchFavTooltip=1; __gsas=ID=e1e467afc7bf6631:T=1672918010:S=ALNI_MbpmOe0b9JBa5bGJmEYB2D7o3PzzQ; laquesis=decision-206@a#deluareb-2066@b#er-1999@c#erm-871@b#euads-3574@a#euads-3583@a#jobs-3834@b#jobs-4078@a#jobs-4633@b#jobs-4643@a#jobs-4757@a#jobs-4782@b#mart-649@a#oesx-2305@a#oesx-2313@b#oesx-2321@b#oesx-2322@c#oeu2u-2590@b#olxeu-40144@a; laquesisff=aut-716#buy-2279#buy-2811#decision-657#euonb-114#euonb-48#grw-124#kuna-307#oesx-1437#oesx-1643#oesx-645#oesx-867#olxeu-29763#srt-1289#srt-1346#srt-1434#srt-1593#srt-1758#srt-477#srt-479#srt-682; lqstatus=1672992323; deviceGUID=689af193-441a-4dfc-9cb0-ecf432e7edc0; a_access_token=0f45f62996e91b197375d0a9b55abb1760d194e9; a_refresh_token=4ed6e87a2cb4b74b1e4e5b389b7985da2576462c; a_grant_type=device; user_business_status=private; from_detail=0; x-device-id=g8obu9o0ljshxg%3A1672991434673; user_id=402139249; _ga=GA1.1.1008601252.1673861517; refresh_token=00835b93caa6706c340eb61f58d834133c9b692e; _ga_W0P4NSQQ21=GS1.1.1678110482.6.1.1678110520.0.0.0; __utma=250720985.1006934100.1672917976.1679075172.1680887543.10; __utmz=250720985.1680887543.10.5.utmcsr=google|utmccn=(organic)|utmcmd=organic|utmctr=(not%20provided); lang=uk; access_token=ccbeda69abccd020feeb186dc993d98667a8613f; sbjs_migrations=1418474375998%3D1; sbjs_current_add=fd%3D2023-04-10%2023%3A07%3A02%7C%7C%7Cep%3Dhttps%3A%2F%2Fwww.olx.ua%2Fd%2Fuk%2Felektronika%2Fkompyutery-i-komplektuyuschie%2Fkomplektuyuschie-i-aksesuary%2Fpol%2F%3Fcurrency%3DUAH%26page%3D13%26search%255Bfilter_enum_subcategory%255D%255B0%255D%3Dvideokarty%26search%255Border%255D%3Dcreated_at%253Adesc%26view%3Dlist%7C%7C%7Crf%3D%28none%29; sbjs_first_add=fd%3D2023-04-10%2023%3A07%3A02%7C%7C%7Cep%3Dhttps%3A%2F%2Fwww.olx.ua%2Fd%2Fuk%2Felektronika%2Fkompyutery-i-komplektuyuschie%2Fkomplektuyuschie-i-aksesuary%2Fpol%2F%3Fcurrency%3DUAH%26page%3D13%26search%255Bfilter_enum_subcategory%255D%255B0%255D%3Dvideokarty%26search%255Border%255D%3Dcreated_at%253Adesc%26view%3Dlist%7C%7C%7Crf%3D%28none%29; sbjs_current=typ%3Dtypein%7C%7C%7Csrc%3D%28direct%29%7C%7C%7Cmdm%3D%28none%29%7C%7C%7Ccmp%3D%28none%29%7C%7C%7Ccnt%3D%28none%29%7C%7C%7Ctrm%3D%28none%29; sbjs_first=typ%3Dtypein%7C%7C%7Csrc%3D%28direct%29%7C%7C%7Cmdm%3D%28none%29%7C%7C%7Ccmp%3D%28none%29%7C%7C%7Ccnt%3D%28none%29%7C%7C%7Ctrm%3D%28none%29; sbjs_udata=vst%3D1%7C%7C%7Cuip%3D%28none%29%7C%7C%7Cuag%3DMozilla%2F5.0%20%28Windows%20NT%2010.0%3B%20Win64%3B%20x64%29%20AppleWebKit%2F537.36%20%28KHTML%2C%20like%20Gecko%29%20Chrome%2F111.0.0.0%20Safari%2F537.36%20OPR%2F97.0.0.0; newrelic_cdn_name=CF; dfp_segment=%5B%5D; PHPSESSID=4lkd9smckc7soaddn6b0m2nn1b; __gpi=UID=00000b9de9ad2fc4:T=1672917975:RT=1681157228:S=ALNI_MbVusxLpPSJCNAyoPQGIJA6rzV7DQ; session_start_date=1681159248844; sbjs_session=pgs%3D6%7C%7C%7Ccpg%3Dhttps%3A%2F%2Fwww.olx.ua%2Fd%2Fuk%2Felektronika%2Fkompyutery-i-komplektuyuschie%2Fkomplektuyuschie-i-aksesuary%2F%3Fcurrency%3DUAH%26search%255Border%255D%3Dcreated_at%3Adesc%26search%255Bfilter_enum_subcategory%255D%255B0%255D%3Dvideokarty%26view%3Dlist',
            'referer': 'https://www.olx.ua/d/uk/elektronika/kompyutery-i-komplektuyuschie/komplektuyuschie-i-aksesuary/?currency=UAH&search%5Border%5D=created_at:desc&search%5Bfilter_enum_subcategory%5D%5B0%5D=videokarty&view=list',
            'sec-ch-ua': '"Not?A_Brand";v="99", "Opera";v="97", "Chromium";v="111"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'sec-fetch-dest': 'empty',
            'sec-fetch-mode': 'cors',
            'sec-fetch-site': 'same-origin',
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/111.0.0.0 Safari/537.36 OPR/97.0.0.0',
            'x-client': 'DESKTOP',
            'x-device-id': '689af193-441a-4dfc-9cb0-ecf432e7edc0',
            'x-fingerprint': 'fbdc4f53959cdb4ade2472495277974c5c6324e638a283d8d11dae88fc3236ac3fef60c9cf99daee3fef60c9cf99daeefaed307981c3a6ca4e1f7a2acddfea3321f2c817b2cc220c3fef60c9cf99daee3fef60c9cf99daee4e1f7a2acddfea33b497a357830277b800d2d70001a8f455308e012c59cf7bdd93ba89d2bc096e295c6324e638a283d85a1778be62509b00e8380ebb100f22a65838a5ba78a32a1b00e97cab1dc631e2f6e9813aa2a88b6b24488822f5af18b0aa6925b0558d723aa2dab357225d0511272ec8e83d084c3e6255da10575393646255da105753936400ab77cc9433c49756b16d11aecc8186428983c35d9b74ca131d7dd9eb490fb0fb7dd9ca7e63276c82beb9cbe2697dfd7fce0dd7835f6d117bf9fa6fdad87a6df0e600f8bcbb93ea00dd0ba2190d70f74348c250bb7bd9817cf587ca35748f297cf587ca35748f297cf587ca35748f297cf587ca35748f297cf587ca35748f297cf587ca35748f297cf587ca35748f297cf587ca35748f297cf587ca35748f297cf587ca35748f29e4788261c6c83237',
            'x-platform-type': 'mobile-html5',
        }
        count = i * 40
        params = {
            'offset': f'{count}',
            'limit': '40',
            'category_id': '458',
            'currency': 'UAH',
            'sort_by': 'created_at:desc',
            'filter_enum_subcategory[0]': 'videokarty',
            'filter_refiners': 'spell_checker',
        }
        response = requests.get('https://www.olx.ua/api/v1/offers/', params=params, headers=headers)
        data = response.json()
        get_offers(data)
        self_link = data["links"]["self"]["href"]
        print(self_link)


def get_offer(item):
    pass
    offer = {}

    offer["url"] = item["url"]
    offer["ID_advert"] = item["id"]
    offer["title"] = item["title"]
    offer["description"] = item["description"].replace('<br />', '')
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
    sleep(4)
    book = openpyxl.load_workbook("database.xlsx")
    sheet = book["data"]

    offer = get_offer(item)

    def find_offer_id(offer):
        for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
            for cell in row:
                if str(offer["ID_advert"]) == cell.value:
                    return cell.coordinate
        # telegram_message(offer)
        return None

    coordinates = find_offer_id(offer)

    if coordinates is not None:

        if str(offer["price"]) != sheet[f'E{coordinates[1:]}'].value:
            sheet[f'C{coordinates[1:]}'].value = offer['title']
            sheet[f'D{coordinates[1:]}'].value = offer["description"]
            sheet[f'E{coordinates[1:]}'].value = str(offer["price"])
            sheet[f'F{coordinates[1:]}'].value = offer["previous_price"]
            sheet[f'G{coordinates[1:]}'].value = offer["last_refresh_time"]
            sheet[f'H{coordinates[1:]}'].value = offer["created_time"]
            sheet[f'I{coordinates[1:]}'].value = offer["city"]
            sheet[f'J{coordinates[1:]}'].value = offer["region"]
            # telegram_message(offer)
            book.save("database.xlsx")

    else:
        sheet.append([offer['ID_advert'], offer['url'], offer['title'], offer["description"], offer["price"], offer["previous_price"], offer["last_refresh_time"], offer["created_time"], offer["city"], offer["region"]])
        book.save("database.xlsx")



def telegram_message(offer):
    chat_id = '-1001696601088'
    message_text = f'О ЧЕТА НОВОЕ : \n Заголовок : \n {offer["title"]} \n Описание : \n {offer["description"]} \n Цена : \n {offer["price"]} \n Предыдущая цена \n {offer["previous_price"]} \n Последнее время изменения \n {offer["last_refresh_time"]} \n Время создания : \n {offer["created_time"]} \n Город : \n {offer["city"]} \n Область : \n {offer["region"]} \n Ссылка : \n{offer["url"]}'
    bot.send_message(chat_id=chat_id, text=message_text)



def get_offers(data):
    offers = []

    for item in data["data"]:
        check_database(item)
        # break

    return offers

keep_alive()
def main():
    get_json()


if __name__ == '__main__':
    main()
