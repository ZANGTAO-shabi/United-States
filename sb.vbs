msgbox"你好啊靳蒙蒙",,"欢迎欢迎"
choose=inputbox("那么，你想使用什么工具呢？"&vbCrLf&"计算器→输入1"&vbCrLf&"圆面积计算→输入2"&vbCrLf&"

圆周长计算→输入3")



if choose=1 then
dim a,m,b,x
a=CDbl(inputbox("请输入第一个数"))
m=inputbox("请输入运算符:+-*/")
b=CDbl(Inputbox("请输入第二个数"))
if m="+" then x=a+b
if m="-" then x=a-b
if m="* "then x=a*b
if m="/ "then x=a/b
msgbox "计算结果是："&x&"。",,"计算完毕"
end if



if choose=2 then
dim r,d,pi,s
pi=3.1415926535
r=inputbox("请输入圆的半径，若要输入直径请留空继续")
if r="" then
d=inputbox("请输入圆的直径")
s=pi*(d/2)*(d/2)
msgbox "您输入的直径为:"&d&"，"&"圆的面积为："&s&"。",,"计算完毕（π按3.1415926535计算）"
else
s=r*r*pi
msgbox "您输入的半径为:"&r&"，"&"圆的面积为："&s&"。",,"计算完毕（π按3.1415926535计算）"
end if



if choose=3 then
pi=3.1415926535
r=inputbox("请输入圆的半径，若要输入直径请留空继续")
if r="" then
d=inputbox("请输入圆的直径")
c=pi*d
msgbox "您输入的直径为:"&d&"，"&"圆的面积为："&c&"。",,"计算完毕（π按3.1415926535计算）"
else
c=pi*r*2
msgbox "您输入的半径为:"&r&"，"&"圆的面积为："&c&"。",,"计算完毕（π按3.1415926535计算）" 
End If
