import smtplib
from email.mime.multipart import MIMEMultipart
import json
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText

def send_gmail(mail_from, mail_to, mail_subject, mail_body):

    settings_file = open('settings.json','r')
    settings_data = json.load(settings_file)

    """ メッセージのオブジェクト """
    msg = MIMEMultipart()
    msg['Subject'] = mail_subject
    msg['From'] = mail_from
    msg['To'] = mail_to

    # エラーキャッチ
    try:
        """ SMTPメールサーバーに接続 """
        smtpobj = smtplib.SMTP('smtp.gmail.com', 587)  # SMTPオブジェクトを作成。smtp.gmail.comのSMTPサーバーの587番ポートを設定。
        smtpobj.ehlo()                                 # SMTPサーバとの接続を確立
        smtpobj.starttls()                             # TLS暗号化通信開始
        gmail_addr = settings_data['gmail_addr']       # Googleアカウント(このアドレスをFromにして送られるっぽい)
        app_passwd = settings_data['app_passwd']       # アプリパスワード
        smtpobj.login(gmail_addr, app_passwd)          # SMTPサーバーへログイン

        body = "test file from tei"
        # attach the body with the msg instance
        msg.attach(MIMEText(body, "plain"))

        # ファイルを添付
        with open(r"./salary/" + mail_to.split('@')[0] + ".pdf", "rb") as f:
            attachment = MIMEApplication(f.read())

        attachment.add_header("Content-Disposition", "attachment", filename=r"./salary/" + mail_to.split('@')[0] + ".pdf")
        msg.attach(attachment)
        """ メール送信 """
        smtpobj.send_message(msg)

        #smtpobj.sendmail(mail_from, mail_to, msg.as_string())

        """ SMTPサーバーとの接続解除 """
        smtpobj.quit()

    except Exception as e:
        print(e)
    
    return "メール送信完了"


mail_list = ['test1@suzukiaudio.co.jp',
             'test2@suzukiaudio.co.jp',
             'test3@suzukiaudio.co.jp',
             'test4@suzukiaudio.co.jp',
             'test5@suzukiaudio.co.jp']


# 直接起動の場合はこちらの関数を実行
if __name__== "__main__":
    for mail_to in mail_list:
            """ メール設定 """
            mail_from = "kobevsnash@gmail.com"       # 送信元アドレス
            mail_to = mail_to                        # 送信先アドレス(To)
            mail_subject = "件名"                    # メール件名
            mail_body = "本文"                       # メール本文

            """ send_gmail関数実行 """
            result = send_gmail(mail_from, mail_to, mail_subject, mail_body)
            print(result, mail_to)

    
