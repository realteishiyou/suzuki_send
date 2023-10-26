import smtplib
from email.mime.text import MIMEText
import json

def send_gmail(mail_from, mail_to, mail_subject, mail_body):

    settings_file = open('settings.json','r')
    settings_data = json.load(settings_file)

    """ メッセージのオブジェクト """
    msg = MIMEText(mail_body, "plain", "utf-8")
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

        """ メール送信 """
        smtpobj.sendmail(mail_from, mail_to, msg.as_string())

        """ SMTPサーバーとの接続解除 """
        smtpobj.quit()

    except Exception as e:
        print(e)
    
    return "メール送信完了"


# 直接起動の場合はこちらの関数を実行
if __name__== "__main__":

    """ メール設定 """
    mail_from = "kobevsnash@gmail.com"       # 送信元アドレス
    mail_to = "dingzy@fuji.waseda.jp"         # 送信先アドレス(To)
    mail_subject = "件名"                   # メール件名
    mail_body = "本文"                      # メール本文

    """ send_gmail関数実行 """
    result = send_gmail(mail_from, mail_to, mail_subject, mail_body)
    print(result)
