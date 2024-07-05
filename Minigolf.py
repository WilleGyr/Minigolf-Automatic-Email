import gspread
import os
from email.message import EmailMessage
import ssl
from email.utils import formataddr
import smtplib
import datetime
import base64
import pandas as pd
import matplotlib.pyplot as plt

today = datetime.datetime.today()

with open ('C:/Users/willi/Desktop/Scripts/GoogleSheets/Minigolf/LastWeek.txt', 'r') as f:
    LastWeek = f.readlines()
    f.close()

Mail = ["","","","","","","",""]

def CheckLastWeek():
    # Willy rekord
    if float(LastWeek[2]) > WScore:
        Mail[0] = f"Willy slog sitt rekord på <b>{str(LastWeek[2].strip())}</b> och fick <b>{WScore}</b>"
    else:
        Mail[0] = f"Willys rekord är fortfarande <b>{str(LastWeek[2].strip())}</b>"
    
    # Willy snitt
    if float(LastWeek[5].replace(',', '.').strip()) == WSnitt:
        Mail[1] = f"Willys snitt är fortfarande <b>{str(LastWeek[5].strip())}</b>"
    if float(LastWeek[5].replace(',', '.').strip()) > WSnitt:
        Mail[1] = f"Willy minskade sitt snitt från <b>{str(LastWeek[5].strip())}</b> till <b>{WSnitt}</b>"
    if float(LastWeek[5].replace(',', '.').strip()) < WSnitt:
        Mail[1] = f"Willy ökade sitt snitt från <b>{str(LastWeek[5].strip())}</b> till <b>{WSnitt}</b>"
    
    # Adri rekord
    if float(LastWeek[3]) > AScore:
        Mail[2] = f"Adri slog sitt rekord på <b>{str(LastWeek[3].strip())}</b> och fick <b>{AScore}</b>"
    else:
        Mail[2] = f"Adris rekord är fortfarande <b>{str(LastWeek[3].strip())}</b>"

    # Adri snitt
    if float(LastWeek[6].replace(',', '.').strip()) == ASnitt:
        Mail[3] = f"Adris snitt är fortfarande <b>{str(LastWeek[6].strip())}</b>"
    if float(LastWeek[6].replace(',', '.').strip()) > ASnitt:
        Mail[3] = f"Adri minskade sitt snitt från <b>{str(LastWeek[3].strip())}</b> till <b>{ASnitt}</b>"
    if float(LastWeek[6].replace(',', '.').strip()) < ASnitt:
        Mail[3] = f"Adri ökade sitt snitt från <b>{str(LastWeek[3].strip())}</b> till <b>{ASnitt}</b>"

    # Dennis rekord
    if float(LastWeek[4]) > DScore:
        Mail[4] = f"Dennis slog sitt rekord på <b>{str(LastWeek[4].strip())}</b> och fick <b>{DScore}</b>"
    else:
        Mail[4] = f"Dennis rekord är fortfarande <b>{str(LastWeek[4].strip())}</b>"
    
    # Dennis snitt
    if float(LastWeek[7].replace(',', '.').strip()) == DSnitt:
        Mail[5] = f"Dennis snitt är fortfarande <b>{str(LastWeek[7].strip())}</b>"
    if float(LastWeek[7].replace(',', '.').strip()) > DSnitt:
        Mail[5] = f"Dennis minskade sitt snitt från <b>{str(LastWeek[7].strip())}</b> till <b>{DSnitt}</b>"
    if float(LastWeek[7].replace(',', '.').strip()) < DSnitt:
        Mail[5] = f"Dennis ökade sitt snitt från <b>{str(LastWeek[7].strip())}</b> till <b>{DSnitt}</b>"

    if float(LastWeek[8].replace(',', '.').strip()) == TotaltSnitt:
        Mail[6] = f"Ert gemensamma snitt är fortfarande <b>{str(LastWeek[8].strip())}</b>"
    if float(LastWeek[8].replace(',', '.').strip()) > TotaltSnitt:
        Mail[6] = f"Ert gemensamma snitt minskade från <b>{str(LastWeek[8].strip())}</b> till <b>{TotaltSnitt}</b>"
    if float(LastWeek[8].replace(',', '.').strip()) < TotaltSnitt:
        Mail[6] = f"Ert gemensamma snitt ökade från <b>{str(LastWeek[8].strip())}</b> till <b>{TotaltSnitt}</b>"
    

sa = gspread.service_account()
sh = sa.open("Minigolf scoreboard")
MOS = sh.worksheet("MOS")
TOTAL = sh.worksheet("Total")

def PlotSnitt():
    SnittLista = MOS.get('U25:X42')

    # Update DataFrame creation to include all columns
    df = pd.DataFrame(SnittLista, columns=['Hål', 'Willy', 'Adri', 'Dennis'])

    # Convert all necessary columns to the correct data type
    df['Hål'] = df['Hål'].astype(int)
    df['Willy'] = df['Willy'].str.replace(',', '.').astype(float)
    df['Adri'] = df['Adri'].str.replace(',', '.').astype(float)
    df['Dennis'] = df['Dennis'].str.replace(',', '.').astype(float)

    # Set the figure size to 468x264 pixels at 300 DPI
    plt.figure(figsize=(468*3 / 300, 264*3 / 300), dpi=300)

    # Plot each of the three x value columns against the hole number with specified colors
    plt.plot(df['Hål'], df['Willy'], label='Willy', color='blue')
    plt.plot(df['Hål'], df['Adri'], label='Adri', color='green')
    plt.plot(df['Hål'], df['Dennis'], label='Dennis', color='orange')

    # Adding labels and legend
    plt.legend()

    # Set the starting y value to 1
    plt.ylim(bottom=1)

    # Ensure every hole number is shown on the x-axis
    plt.xticks(df['Hål'])

    # Save the figure to a file
    plt.savefig('C:/Users/willi/Desktop/Scripts/GoogleSheets/Minigolf/SnittGraf.png', dpi=300, bbox_inches='tight')

def ImageToBase64():
    with open('C:/Users/willi/Desktop/Scripts/GoogleSheets/Minigolf/SnittGraf.png', 'rb') as img:
        base64_image = base64.b64encode(img.read()).decode()
    return base64_image

PlotSnitt()
base64_image = ImageToBase64()

Topplista = MOS.get('C31:E40')

def ReplaceComma():
    for i in LastWeek:
        i = i.replace(',', '.').strip()

WScore = MOS.acell('C25').value
WScore = int(WScore.replace(',', '.'))
AScore = MOS.acell('D25').value
AScore = int(AScore.replace(',', '.'))
DScore = MOS.acell('E25').value
DScore = int(DScore.replace(',', '.'))

WSnitt = MOS.acell('V43').value
WSnitt = float(WSnitt.replace(',', '.'))
ASnitt = MOS.acell('W43').value
ASnitt = float(ASnitt.replace(',', '.'))
DSnitt = MOS.acell('X43').value
DSnitt = float(DSnitt.replace(',', '.'))

TotaltSnitt = MOS.acell('M43').value
TotaltSnitt = float(TotaltSnitt.replace(',', '.'))

AntalHIO = MOS.acell('D43').value
AntalHIO = int(AntalHIO)

AntalSlag = MOS.acell('K51').value
AntalSlag = int(AntalSlag.replace('\xa0', '').replace(' ', ''))

CurrentCell = MOS.cell(21, int(LastWeek[11]))
OldCell = CurrentCell

while CurrentCell.value != None:
    #print(CurrentCell.value)
    CurrentCell = MOS.cell(CurrentCell.row, CurrentCell.col + 1)

def GetAverage(startcell, stopcell):
    start = startcell.row, startcell.col
    stop = stopcell.row, stopcell.col
    sum = 0
    count = 0
    for row in range(start[0], stop[0] + 1):
        for col in range(start[1], stop[1] + 1):
            cell = MOS.cell(row, col)
            if cell.value != None:
                sum += int(cell.value)
                count += 1
    if count > 0:
        return sum / count
    else:
        return 0  # or return 0, depending on the desired behavior when there are no values

VeckoSnitt = GetAverage(OldCell, CurrentCell)
VeckoSnitt = round(VeckoSnitt, 1)

#print(Topplista[1][2])

Rundor = int(TOTAL.acell('H4').value)

#ReplaceComma()
CheckLastWeek()


def sendMail(EMAIL_RECIEVER):
    EMAIL_ADDRESS = 'autosendminigolfgyrulf@gmail.com'
    EMAIL_PASSWORD = 'nwgy ludp gxcp lazc'
    
    msg = EmailMessage()
    msg['Subject'] = f'Summering Minigolf vecka {str(today.isocalendar()[1])}'
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = EMAIL_RECIEVER
    msg.set_content(f'''
        <!DOCTYPE html>
        <html>
        <head>
            <link rel="stylesheet" type="text/css" hs-webfonts="true" href="https://fonts.googleapis.com/css?family=Lato|Lato:i,b,bi">
            <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <style type="text/css">
                body {{
                    width: 100%;
                    font-family: Lato, sans-serif;
                    font-size: 18px;
                    background-color: #F5F8FA;
                    margin: 0;
                    padding: 0;
                }}
                #email {{
                    margin: auto;
                    width: 600px;
                    background-color: #fff;
                    border-radius: 8px;
                    overflow: hidden;
                    box-shadow: 0 4px 8px rgba(0,0,0,0.1);
                }}
                .header {{
                    background: url('https://www.visitorebro.se/wp-content/uploads/2021/06/Minigolf_Alka%CC%88rretec411b53fcf24e868bb9fc77facbfb454e76a4b67965b401739891f156674d88-1170x660.jpg') no-repeat center center;
                    background-size: cover;
                    color: white;
                    text-align: center;
                    padding: 40px 20px;
                    position: relative;
                }}
                .header h1 {{
                    font-size: 48px;
                    margin: 0;
                    position: relative;
                    z-index: 2;
                }}
                .header:before {{
                    content: '';
                    position: absolute;
                    top: 0;
                    right: 0;
                    bottom: 0;
                    left: 0;
                    background: rgba(0, 0, 0, 0.5);
                    z-index: 1;
                }}
                .content {{
                    padding: 20px 30px;
                }}
                .content h2 {{
                    font-size: 28px;
                    font-weight: 900;
                    margin: 20px 0 10px;
                    color: #333;
                }}
                .content p {{
                    font-weight: 100;
                    line-height: 1.6;
                    margin: 10px 0;
                    color: #666;
                }}
                .section {{
                    padding: 20px;
                    margin-bottom: 20px;
                    border: 1px solid #ddd;
                    border-radius: 8px;
                    background-color: #fafafa;
                }}
                .table-section {{
                    padding: 20px;
                    margin-bottom: 20px;
                    border: 1px solid #ddd;
                    border-radius: 8px;
                    background-color: #fafafa;
                }}
                .table-section table {{
                    width: 100%;
                    border-collapse: collapse;
                }}
                .table-section th, .table-section td {{
                    padding: 8px;
                    text-align: center;
                    border: 1px solid #ddd;
                }}
                .table-section th {{
                    background-color: #f1f1f1;
                }}
                .footer {{
                    text-align: center;
                    color: #999;
                    padding: 20px;
                    font-size: 14px;
                    background-color: #f1f1f1;
                    border-top: 1px solid #ddd;
                }}
            </style>
        </head>
        <body>
        <div id="email">
            <div class="header">
                <h1>Summering Minigolf V.{str(today.isocalendar()[1])}</h1>
            </div>
            <div class="content">
                <div class="section">
                    <h2>Totalt</h2>
                    <p>
                        Ni spelade <b>{int(Rundor) - int(LastWeek[0])}</b> rundor<br>
                        Ni slog <b>{str(AntalSlag - int(LastWeek[10]))}</b> slag<br>
                        Ni fick <b>{str(AntalHIO - int(LastWeek[9]))}</b> Hole-In-Ones<br>
                        Ni hade ett snitt på <b>{str(VeckoSnitt)}</b> per runda<br>
                        Ert nuvarande rekord är <b>{str(Topplista[0][2])}</b><br>
                        {str(Mail[6])}
                    </p>
                </div>
                <div class="section">
                    <h2>Willy</h2>
                    <p>
                        {str(Mail[0])}<br>
                        {str(Mail[1])}<br>
                    </p>
                </div>
                <div class="section">
                    <h2>Adri</h2>
                    <p>
                        {str(Mail[2])}<br>
                        {str(Mail[3])}<br>
                    </p>
                </div>
                <div class="section">
                    <h2>Dennis</h2>
                    <p>
                        {str(Mail[4])}<br>
                        {str(Mail[5])}
                    </p>
                </div>
                <div class="section">
                    <h2>Snitt Per Hål</h2>
                    <p>
                        <img src="data:image/png;base64,{base64_image}" alt="SnittGraf" title="SnittGraf" height=280 width=468>
                    </p>
                </div>
                <div class="table-section">
                    <h2>Topplistan</h2>
                    <table>
                        <thead>
                            <tr>
                                <th>Rank</th>
                                <th>Name</th>
                                <th>Score</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td>1</td>
                                <td>{Topplista[0][1]}</td>
                                <td>{Topplista[0][2]}</td>
                            </tr>
                            <tr>
                                <td>2</td>
                                <td>{Topplista[1][1]}</td>
                                <td>{Topplista[1][2]}</td>
                            </tr>
                            <tr>
                                <td>3</td>
                                <td>{Topplista[2][1]}</td>
                                <td>{Topplista[2][2]}</td>
                            </tr>
                            <tr>
                                <td>4</td>
                                <td>{Topplista[3][1]}</td>
                                <td>{Topplista[3][2]}</td>
                            </tr>
                            <tr>
                                <td>5</td>
                                <td>{Topplista[4][1]}</td>
                                <td>{Topplista[4][2]}</td>
                            </tr>
                            <tr>
                                <td>6</td>
                                <td>{Topplista[5][1]}</td>
                                <td>{Topplista[5][2]}</td>
                            </tr>
                            <tr>
                                <td>7</td>
                                <td>{Topplista[6][1]}</td>
                                <td>{Topplista[6][2]}</td>
                            </tr>
                            <tr>
                                <td>8</td>
                                <td>{Topplista[7][1]}</td>
                                <td>{Topplista[7][2]}</td>
                            </tr>
                            <tr>
                                <td>9</td>
                                <td>{Topplista[8][1]}</td>
                                <td>{Topplista[8][2]}</td>
                            </tr>
                            <tr>
                                <td>10</td>
                                <td>{Topplista[9][1]}</td>
                                <td>{Topplista[9][2]}</td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>
            <div class="footer">
                &copy; 2024 The Racist Lads. All rights reserved.
            </div>
        </div>
        </body>
        </html>
    ''', subtype='html')
    
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        smtp.send_message(msg)
        print("Successfully sent the mail.")

with open ('C:/Users/willi/Desktop/Scripts/GoogleSheets/Minigolf/LastWeek.txt', 'w') as f:
    f.write(str(Rundor))
    f.write('\n')
    f.write(str(Topplista[0][2]))
    f.write('\n')
    f.write(str(WScore))
    f.write('\n')
    f.write(str(AScore))
    f.write('\n')
    f.write(str(DScore))
    f.write('\n')
    f.write(str(WSnitt))
    f.write('\n')
    f.write(str(ASnitt))
    f.write('\n')
    f.write(str(DSnitt))
    f.write('\n')
    f.write(str(TotaltSnitt))
    f.write('\n')
    f.write(str(AntalHIO))
    f.write('\n')
    f.write(str(AntalSlag))
    f.write("\n")
    f.write(str(CurrentCell.col))
    f.close()

sendMail('william.gyrulf@hotmail.com')
sendMail('adriandushi@outlook.com')
sendMail('dennis.tollofsen04@gmail.com')