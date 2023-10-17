# -*- coding: UTF-8 -*-
import requests as req
import json,sys,time
#先注册azure应用,确保应用有以下权限:
#files:	Files.Read.All、Files.ReadWrite.All、Sites.Read.All、Sites.ReadWrite.All
#user:	User.Read.All、User.ReadWrite.All、Directory.Read.All、Directory.ReadWrite.All
#mail:  Mail.Read、Mail.ReadWrite、MailboxSettings.Read、MailboxSettings.ReadWrite
#注册后一定要再点代表xxx授予管理员同意,否则outlook api无法调用






path=sys.path[0]+r'/1.txt'
num1 = 0

def gettoken(refresh_token):
    headers={'Content-Type':'application/x-www-form-urlencoded'
            }
    data={'grant_type': 'refresh_token',
          'refresh_token': refresh_token,
          'client_id':id,
          'client_secret':secret,
          'redirect_uri':'https://musummer.github.io/GetAutoApiToken/'
         }
    html = req.post('https://login.microsoftonline.com/common/oauth2/v2.0/token',data=data,headers=headers)
    jsontxt = json.loads(html.text)
    refresh_token = “0.ASsAbXOEstenk0u6g41mjMc15IZMP0riKihNt6tevifujVHCAJo.AgABAAEAAAAtyolDObpQQ5VtlI4uGjEPAgDs_wUA9P_SyvAbJlwQfq1thGOYswuTL9uRpUW4npSy46HtZjZ3VpCvLVJCj3UogaH3tjeCbMkNx3Ai5P8hqcIuZYqvE8bk73bZxWgxLROEcSDD71tVT0VandY5_-UqFYwghx-ne9OMOQu94WESY2kPExxFGisu24VLO1214x4xqoMhjN_CrRDIoyPDUCOhnAotZddKr6u1YayvMXLJAsMpk4B2oYm67fs0Y5yy5b0d9ZLm26xvMX647Fkg4I1epZJDi6p2acnK9lKdywtOF0eCztpAIlhS2bwIS55UE3LzBDq6ZVH4ERwYF8hLryS4mGwT5E9GyQsxanKVO80tcVaOzok__cfVWNd6oHTTfOJBsqAn8bsAE0mxMHMjKWcfGUDPxlG4dQd0QUYxrjj5PyjgZhoVY3Acomz511nEbOgPNjqr4k5Z7n2jU7QxJzILSmark31e5v3ZAVHIS0ZuTOKrAiGmEHXqrkD1MUrrzdAFrxchHoxuRqH6QBIMxwr5oOegkgoEAWJk1QF8Uf4gAXAyjeApW6BQmTiPXEg4fV37RFbpEtDK-k75nFC1DLwSf5-sZflVcyn3FP89KwYX4YrqGxMzdFC_kOEvNfni-v5cekJQgXUzS-63GfPVaEq1yvXP8kca7Zg-VUl5748iCbcOLxTXehapFgpvDPjGruRjAJYinDB9mb8a7weVPO5Tnj9GbrQhkXUZTT_EEIilwhMVWuWc0kinkP5z9zIxZeTDFoUk3473Bw94tY1c7IqTEwB7EJdtnpshKUKfF_btJpkBDLykvNo”
    access_token = jsontxt['access_token']
    with open(path, 'w+') as f:
        f.write(refresh_token)
    return access_token
def main():
    fo = open(path, "r+")
    refresh_token = fo.read()
    fo.close()
    global num1
    localtime = time.asctime( time.localtime(time.time()) )
    access_token=gettoken(refresh_token)
    headers={
    'Authorization':access_token,
    'Content-Type':'application/json'
    }
    try:
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive/root',headers=headers).status_code == 200:
            num1+=1
            print("1调用成功"+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive',headers=headers).status_code == 200:
            num1+=1
            print("2调用成功"+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/drive/root',headers=headers).status_code == 200:
            num1+=1
            print('3调用成功'+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/users ',headers=headers).status_code == 200:
            num1+=1
            print('4调用成功'+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/messages',headers=headers).status_code == 200:
            num1+=1
            print('5调用成功'+str(num1)+'次')    
        if req.get(r'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules',headers=headers).status_code == 200:
            num1+=1
            print('6调用成功'+str(num1)+'次')    
        if req.get(r'https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages/delta',headers=headers).status_code == 200:
            num1+=1
            print('7调用成功'+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive/root/children',headers=headers).status_code == 200:
            num1+=1
            print('8调用成功'+str(num1)+'次')
        if req.get(r'https://api.powerbi.com/v1.0/myorg/apps',headers=headers).status_code == 200:
            num1+=1
            print('8调用成功'+str(num1)+'次') 
        if req.get(r'https://graph.microsoft.com/v1.0/me/mailFolders',headers=headers).status_code == 200:
            num1+=1
            print('9调用成功'+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/outlook/masterCategories',headers=headers).status_code == 200:
            num1+=1
            print('10调用成功'+str(num1)+'次')
            print('此次运行结束时间为 :', localtime)
    except:
        print("pass")
        pass
for _ in range(9):
    main()
