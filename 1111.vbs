On Error Resume Next

Dim wsh,ye

set wsh=createobject("wscript.shell")

for i=1 to 10

wscript.sleep 100

wsh.AppActivate("�� xx ������")

wsh.sendKeys "^v"

wsh.sendKeys i

wsh.sendKeys "%s"

next

wscript.quit

