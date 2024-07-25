#! python2.7
# cntrl + shift + p to switch interpreters

from twilio.rest import Client

#pip3 install twilio
account_sid = ''
auth_token = ''
client = Client(account_sid, auth_token)

message = client.messages.create(
  from_='+18884106031',
  body='Hello from Twilio',
  to='+13179891988'
)

print(message.sid)