function SetMail() {
    cdoMsg = new ActiveXObject("CDO.Message") // create CDO.Message object
    cdoMsg.From = "sender@163.com" // sender
    cdoMsg.To = "recipient@qq.com" // recipient
    namespace = "http://schemas.microsoft.com/cdo/configuration/" // CDO namespace
    cdoMsg.Configuration.Fields.Item(namespace + "sendusing") = 2 // use network send
    cdoMsg.Configuration.Fields.Item(namespace + "smtpserver") = "smtp.163.com" // SMTP Server
    cdoMsg.Configuration.Fields.Item(namespace + "smtpserverport") = 25 // SMTP port number
    cdoMsg.Configuration.Fields.Item(namespace + "smtpauthenticate") = 1 // Set Basic
    cdoMsg.Configuration.Fields.Item(namespace + "sendusername") = "sender@163.com" // username
    cdoMsg.Configuration.Fields.Item(namespace + "sendpassword") = "SDFSDFGWERFFSGJ" // password
    cdoMsg.Configuration.Fields.Item(namespace + "smtpusessl") = true // use ssl encryption
    cdoMsg.Configuration.Fields.Update() // auto update
}

function SendMail() {
    cdoMsg.Subject = title // title
    cdoMsg.Textbody = cry // content
    //cdoMsg.AddAttachment("1.png") //send files
    cdoMsg.Send() // send mail
    cdoMsg = null //clean
}

function ExCmd() {
    //Get now time
    date = new Date().toLocaleDateString()
    time = new Date().toLocaleTimeString()
    dt = date + " " + time

    //Get PC name
    wn = new ActiveXObject("WScript.Network")
    PcName = wn.computername
    user = wn.username
    ws = new ActiveXObject("WScript.Shell")

    //Output port status
    ws.run(
        "cmd /c del /q /f \"%temp%\\port.tmp\"&&(\
            for /f \"tokens=2,3\" %a in ('netstat -ano^|find \"3389\"^|find \"ESTABLISHED\"') do (\
                echo %a %b >>\"%temp%\\port.tmp\"\
            )\
        )\
    ",0,true)

    //Get port status
    fso = new ActiveXObject("Scripting.FileSystemObject")
    temp = fso.GetSpecialFolder(2) + '\\port.tmp'
    file = fso.OpenTextFile(temp,1)
    buffer = 0
    while (! file.AtEndOfStream) {
        str = file.ReadLine()
        str = str.split(':')[1]
        ip = str.split(' ')[1]
        port = str.split(' ')[0]
        if (port == 3389) {
            LogStat = 'Remote Login'
            cry = 'IP:'+ ip + ' Port:' + port + ' User:' + user + ',Log in to the host at ' + dt + '.'
        } else {
            buffer++
        }
    }
    file.Close()
    fso.DeleteFile(temp)
    if (i > 0) {
        LogStat = 'Local Login'
        cry = 'User:' + user + ',Log in to the host at ' + dt + '.'
    }
    title = '[' + PcName + '] ' + LogStat
}
SetMail()
ExCmd()
SendMail()
