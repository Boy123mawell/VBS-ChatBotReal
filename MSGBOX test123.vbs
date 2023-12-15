dim msg, sapi
msg=inputbox("Type your text inside the input bot, only can use English.","Speaking AI")
set sapi=Createobject("sapi.spvoice")
sapi.speak msg