# django-word2pdf

## 开发环境

+ Windows 10
+ python2.7

## 安装程序依赖

```
pip install -r .\requirements.txt
```

## 程序运行

```
.\run.bat
```

## 开发中遇到的典型问题

### pywintypes.com_error: (-2147221005, '\xce\xde\xd0\xa7\xb5\xc4\xc0\xe0\xd7\xd6\xb7\xfb\xb4\xae', None, None)

 原因是缺少程序：[CAPICOM](https://www.microsoft.com/en-us/download/details.aspx?id=25281), 但是下载的capicon是32位的，需要手工注册到64位下。
 
 解决过程如下：

+ 安装软件：[CAPICOM](https://www.microsoft.com/en-us/download/details.aspx?id=25281)
+ 到安装目录下找到capicom.dll, 并复制到SysWOW64下
+ 执行如下命令进行注册
 
```
cd C:\Windows\SysWow64 && regsvr32.exe capicon.dll
```

### pywintypes.com_error: (-2147221008, '\xc9\xd0\xce\xb4\xb5\xf7\xd3\xc3 CoInitialize\xa1\xa3', None, None)

原因是多线程编程，调用了win32com模块

解决方式：
+ 引入pythoncom

```
from pythoncom import CoInitialize, CoUninitialize
```

+ 在word对象Dispatch()和Quit()时，分别调用CoInitialize()和CoUnintialize()

```
CoInitialize()
word = Dispatch('word.application')
...
word.Quit(constants.wdDoNotSaveChanges)
CoUnintialize()
```

### pywintypes.com_error -2147352567

原因是文件不支持相对路径的读取

## 其他异常问题的排查

### 文件被询问“是否另存为”：

修改word的Quit()方式为： word.Quit(constants.wdDoNotSaveChanges)

## SC命令

> 需要重启

+ 注册服务

```
.\register.bat
```

+ 删除服务

```
.\unregister.bat
```

+ 启动服务

```
.\startServer.bat
```

+ 关闭服务

```
.\stopServer.bat
```