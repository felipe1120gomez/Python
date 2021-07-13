import json
import time
import os
import xlsxwriter
import PySimpleGUI as sg
from requests import Session
from requests.exceptions import ConnectionError, Timeout, TooManyRedirects

#We obtain the user's data through a GUI.
sg.theme('Dark Black 1')

layout = [[sg.Text('Enter the currency for conversion of values.', justification='center', size=(35,1))],
    [sg.Text('e.g USD or EUR.', justification='center', size=(35,1))],
    [sg.Text(' '*27), sg.InputText(key='-IN-', size=(4, 0), do_not_clear=False)],
    [sg.Text(' '*20), sg.Submit(), sg.Button('Exit')]]

window = sg.Window('Currency for Conversion', layout, resizable=True)

while True:
    event, values = window.read()

    if event in (sg.WIN_CLOSED, 'Exit'):
        break
    if event == sg.WIN_CLOSED or event == 'Exit':
        break

    if event == 'Submit':

        if (len(values['-IN-'])) != 3:
            sg.popup('Invalid input', values['-IN-'])
            continue


        if not values['-IN-'].isalpha():
            sg.popup('Invalid input', values['-IN-'])
            continue


        conversion = values['-IN-'].upper()
        window.close()

window.close()

try:
    sg.popup('Conversion', 'The values will be converted to ' + conversion )
except:
    quit()

layout = [[sg.Text('Enter the symbol of Cryptocurrency and the amount you owned.', justification='center', size=(46,1))],
    [sg.Text('e.g BTC or ETH.', justification='center', size=(46,1))],
    [sg.Text('Cryptocurrency', justification='center', size=(46,1))],
    [sg.Text(' '*36), sg.InputText(key='-IN-', size=(7, 0), do_not_clear=False)],
    [sg.Text(' ', justification='center', size=(46,1))],
    [sg.Text('Amount Owned', justification='center', size=(46,1))],
    [sg.Text('Do not use "," instead use "."', justification='center', size=(46,1))],
    [sg.Text(' '*29), sg.InputText(key='-NUM-', size=(15, 0), do_not_clear=False)],
    [sg.Text('Press Submit to add it to your portfolio, press Save to continue.', justification='center', size=(46,1))],
    [sg.Text(' '*25), sg.Submit(), sg.Button('Save'), sg.Button('Exit')]]

window = sg.Window('Cryptocurrency Portfolio', layout, resizable=True)

my_portfolio = dict()
while True:
    event, values = window.read()

    if event in (sg.WIN_CLOSED, 'Exit'):
        break
    if event == sg.WIN_CLOSED or event == 'Exit':
        break

    if event == 'Submit':

        if not values['-IN-'].isalpha():
            sg.popup('Invalid input', values['-IN-'])
            continue

        if (len(values['-IN-'])) < 2:
            sg.popup('Invalid input', values['-IN-'])
            continue

        try:
            float(values['-NUM-'])

        except:
            sg.popup('Invalid input', values['-NUM-'])
            continue

        my_portfolio[values['-IN-'].upper()] = float(values['-NUM-'])


    if event == 'Save':
        window.close()

        layout = [[sg.Output(size=(60,10), pad=(30,0))],
                 [sg.Text('Press Show to see your portfolio.', justification='center', size=(60,1))],
                 [sg.Text('Press Go to continue.', justification='center', size=(60,1))],
                 [sg.Text(' '*41), sg.Button('Show'), sg.Button('Go'), sg.Button('Exit')]]

        window = sg.Window('Your portfolio.', layout, resizable=True)

        while True:
            event, values = window.read()

            if event in (sg.WIN_CLOSED, 'Exit'):
                break
            if event == sg.WIN_CLOSED or event == 'Exit':
                break

            if event == 'Show':
                for key, value in my_portfolio.items():
                    print('Coin: {}, Amount: {}'.format(key, value), '\n')

            if event == 'Go':
                window.close()

window.close()

try:
    my_coins = list()
    for key, value in my_portfolio.items():
        if key not in my_coins:
            my_coins.append(key)
except:
    quit()

layout = [[sg.Text('Select how many hours you want to update your portfolio information.', justification='center', size=(55,1))],
    [sg.Listbox((0,0.5,1,2,3,4,5,6,7,8), [1], size=(3,10), key='-HR-', pad=((210,210),(20,20)))],
    [sg.Text(' '*41), sg.Submit(), sg.Button('Exit')]]

window = sg.Window('Update settings', layout, resizable=True)

while True:
    event, values = window.read()

    if event in (sg.WIN_CLOSED, 'Exit'):
        break
    if event == sg.WIN_CLOSED or event == 'Exit':
        break

    if event == 'Submit':
        if float(values['-HR-'][0]) > 0.0:
            CHOICE = float(values['-HR-'][0])
            sg.popup('The portfolio will be updated every ' + str(CHOICE) + ' hours.')
            repeat = True
            times = 1
        else:
            sg.popup('The portfolio will only be updated once.')
            CHOICE = 0
            repeat = False
            times = 1

        window.close()

window.close()

try:
    times == 1
except:
    quit()

layout = [[sg.Text('Enter the name for the Excel file.', justification='center', size=(35,1))],
    [sg.Text('A different name every time.', justification='center', size=(35,1))],
    [sg.Text(' '*21), sg.InputText(key='-IN-', size=(12, 0), do_not_clear=False)],
    [sg.Text(' '*20), sg.Submit(), sg.Button('Exit')]]

window = sg.Window('Name for the Excel file.', layout, resizable=True)

while True:
    event, values = window.read()

    if event in (sg.WIN_CLOSED, 'Exit'):
        break
    if event == sg.WIN_CLOSED or event == 'Exit':
        break

    if event == 'Submit':

        if (len(values['-IN-'])) < 1:
            sg.popup('Invalid name', values['-IN-'])
            continue

        FILE_NAME = values['-IN-'] + '.xlsx'

        if os.path.exists(FILE_NAME):
            sg.popup('File already exists', FILE_NAME)
            FILE_NAME = None
            continue

        window.close()

window.close()

try:
    sg.popup('Excel file', 'The name of the Excel file will be ' + FILE_NAME )
except:
    quit()

#The Excel workbook is created.
crypto_workbook = xlsxwriter.Workbook(FILE_NAME)
crypto_sheet = crypto_workbook.add_worksheet('CryptoData')
bold = crypto_workbook.add_format({'bold': True})
money = crypto_workbook.add_format({'num_format': '$#,##0'})
crypto_sheet.set_column('A:M', 22)

crypto_sheet.write('A1', 'NAME', bold)
crypto_sheet.write('B1', 'SYMBOL', bold)
crypto_sheet.write('C1', 'AMOUNT OWNED', bold)
crypto_sheet.write('D1', 'PRICE IN ' + conversion, bold)
crypto_sheet.write('E1', 'OWNED IN ' + conversion, bold)
crypto_sheet.write('F1', 'CHANGE 1H', bold)
crypto_sheet.write('G1', 'CHANGE 24H', bold)
crypto_sheet.write('H1', 'CHANGE 7D', bold)
crypto_sheet.write('I1', 'MARKET CAP', bold)
crypto_sheet.write('J1', 'VOLUME 24H', bold)
crypto_sheet.write('K1', 'CIRCULATING SUPPLY', bold)
crypto_sheet.write('L1', 'LAST UPDATE', bold)
crypto_sheet.write('M1', 'TOTAL PORFOLIO IN ' + conversion, bold)

#The API is called.
seconds = CHOICE * 3600
names = list()
try:
    while repeat or times > 0:

        URL = 'https://pro-api.coinmarketcap.com/v1/cryptocurrency/listings/latest'
        parameters = {
            'start': 1,
            'limit': 1000,#Only the first 1000 coins are taken into account.
            'convert': conversion,
        }
        headers = {
            'Accepts': 'application/json',
            'X-CMC_PRO_API_KEY': 'Paste your API key here',#Paste your API key here.
        }

        session = Session()
        session.headers.update(headers)

        try:
            response = session.get(URL, params = parameters)

        except (ConnectionError, Timeout, TooManyRedirects) as e:
            print(e)
            sg.popup('Connection Error.')
            quit()

        try:
            data = json.loads(response.text)

        except:
            print('==== Failure To Retrieve ====')
            print(response.text)
            sg.popup('Failure To Retrieve.')
            quit()

        if 'data' not in data :
            print('==== The JSON file does not meet the requirements ====')
            sg.popup('The JSON file does not meet the requirements.')
            print(data)
            quit()

        if times > 0:
            row = 1
        else:
            row = len(names) + 2

        portfolio_value = 0

        #We iterate for each currency in the JSON file.
        file = data['data']
        for key in file:
            name = key['name']
            symbol = key['symbol']
            active = key['circulating_supply']
            last = key['last_updated']
            quote = key['quote']
            convert = quote[conversion]
            price = convert['price']
            cap = convert['market_cap']
            hour = convert['percent_change_1h']
            day = convert['percent_change_24h']
            week = convert['percent_change_7d']
            vol = convert['volume_24h']

            #We only add the currencies that are in the user's portfolio.
            for key_1 in my_portfolio.keys():
                if key_1 == symbol:
                    value = float(price) * float(my_portfolio[key_1])

                    crypto_sheet.write(row, 0, name)
                    crypto_sheet.write(row, 1, symbol)
                    crypto_sheet.write(row, 2, my_portfolio[key_1], money)
                    crypto_sheet.write(row, 3, price, money)
                    crypto_sheet.write(row, 4, value, money)
                    crypto_sheet.write(row, 5, str('{:.2f}'.format(hour) + '%'))
                    crypto_sheet.write(row, 6, str('{:.2f}'.format(day) + '%'))
                    crypto_sheet.write(row, 7, str('{:.2f}'.format(week) + '%'))
                    crypto_sheet.write(row, 8, cap, money)
                    crypto_sheet.write(row, 9, vol, money)
                    crypto_sheet.write(row, 10, active, money)
                    crypto_sheet.write(row, 11, last)
                    names.append(symbol)
                    portfolio_value += value
                    row += 1

        crypto_sheet.write(row, 0, '**')
        crypto_sheet.write(row, 12, portfolio_value, money)
        names.append(' ')

        if repeat:

            layout = [[sg.Text('The portfolio has been updated.', justification='center', size=(35,1))],
                [sg.Text('Do you want to end the program?', justification='center', size=(35,1))],
                [sg.Text(' '*23), sg.Button('Yes'), sg.Button('No')]]

            window = sg.Window('Portfolio updated', layout, resizable=True)

            while True:
                event, values = window.read()

                if event == sg.WIN_CLOSED:
                    break

                if event == 'Yes':
                    repeat = False
                    seconds = 1
                    window.close()

                elif event == 'No':
                    sg.popup('The portfolio will be updated once more.')
                    window.close()

            window.close()

        else:
            seconds = 1

        times -= 1
        time.sleep(seconds)

    crypto_workbook.close()
    sg.popup('Open ' + FILE_NAME + ' file.')

except KeyboardInterrupt:
    crypto_workbook.close()
    sg.popup('Program finished by user.')

missing = list()
for coin in my_coins:
    if coin not in names:
        missing.append(coin)

if len(missing) > 0:

    layout = [[sg.Output(size=(60,10), pad=(30,0))],
             [sg.Text('Press Show to see cryptocurrencies not found.', justification='center', size=(60,1))],
             [sg.Text('Press Exit to finish.', justification='center', size=(60,1))],
             [sg.Text(' '*46), sg.Button('Show'), sg.Button('Exit')]]

    window = sg.Window('Cryptocurrencies not found.', layout, resizable=True)

    while True:
        event, values = window.read()

        if event in (sg.WIN_CLOSED, 'Exit'):
            break
        if event == sg.WIN_CLOSED or event == 'Exit':
            break

        if event == 'Show':
            for coin in missing:
                print('Cryptocurrency ' + coin + ' Not found', '\n')

    window.close()
