from __future__ import print_function
import os
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import pandas as pd
import base64
import pickle

# OAuth2.0 Gmail　API用スコープ
# 変更する場合は、token.jsonファイルを削除する
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
dir = os.getcwd()
# アクセストークン取得
def get_token():
    creds = None
    # token.json
    # access/refresh tokenを保存
    # 認可フロー完了時に自動で作成。
    if os.path.exists('{}/creds/token.json'.format(dir)):
        creds = Credentials.from_authorized_user_file(
            'creds/token.json', SCOPES)
    # トークンが存在しない場合
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                '{}/creds/credentials.json'.format(dir), SCOPES)
            creds = flow.run_local_server(port=0)
        # トークンを保存
        with open('{}/creds/token.json'.format(dir), 'w') as token:
            token.write(creds.to_json())
    return creds


# メールリスト取得
def gmail_get_messages_body(service, user_id, query, count):
    message_list = []
    # メッセージの一覧を取得
    messages = service.users().messages()
    msg_list = messages.list(userId=user_id, maxResults=count, q=query).execute()
    message_list = gmail_get_messages_body_content(messages, msg_list, message_list)
    #print(msg_list)
    while True:
        page_token = msg_list['nextPageToken']
        msg_list = messages.list(userId=user_id, maxResults=count, q=query, pageToken=page_token).execute()
        if 'nextPageToken' not in msg_list:
            break
        else:
            message_list = gmail_get_messages_body_content(messages, msg_list, message_list)
    return message_list

def gmail_get_messages_body_content(messages, msg_list, message_list):
    # 各message内容確認
    for message_id in msg_list['messages']:
        # 各メッセージ詳細
        msg = messages.get(userId='me', id=message_id['id']).execute()
        #送り主
        decoded_message = []
        decoded_message.append([
            header["value"]
            for header in msg["payload"]["headers"]
            if header["name"] == "From"
        ][0])
        #タイトル
        decoded_message.append([
            header["value"]
            for header in msg["payload"]["headers"]
            if header["name"] == "Subject"
        ][0])
        # 本文
        if(msg["payload"]["body"]["size"]!=0):
            decoded_message.append(decode_msg(msg["payload"]["body"]["data"]))
        else:
            #メールによっては"parts"属性の中に本文がある場合もある
            decoded_message.append(decode_msg(msg["payload"]["parts"][0]["body"]["data"]))
        message_list.append(decoded_message)
    return message_list

def decode_msg(data):
    decoded_bytes = base64.urlsafe_b64decode(data)
    decoded_msg = decoded_bytes.decode("UTF-8")
    return decoded_msg

def main(query="is:unread",count=10):
    #アクセストークンの取得
    creds = get_token()
    #Gmail API呼び出し
    service = build('gmail', 'v1', credentials=creds)
    message_list = gmail_get_messages_body(service=service,user_id="me",query=query,count = count)
    return message_list

def pickle_dump(obj, path):
    with open(path, mode='wb') as f:
        pickle.dump(obj,f)

def pickle_load(path):
    with open(path, mode='rb') as f:
        data = pickle.load(f)
        return data

if __name__ == '__main__':
    message_list = main()
    pickle_dump(message_list,"{}/mails/gmail.pickle".format(dir))
    df = pd.DataFrame(message_list)
    df.to_csv("{}/mails/gmail.csv".format(dir),index=False,encoding="utf-8-sig")