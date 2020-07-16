# windowsPrinter

windows printer

主要唤醒系统打印机执行打印

支持pdf、word、excel、jpg四种类型的文件打印

用到了jacob.jar(java-COM中间件)，调用win32库实现word、excel 打印，pdf打印使用的为Apache的pdfbox（注：pdfbox打印可能会出现中文字体乱码的问题，只需要把pdfbox和fontbox使用的版本提高即可解决）

项目中将jacob.jar放入了lib文件夹下，解压之后，使用mvn命令上传到本地maven仓库即可
(mvn install:install-file -Dfile=F:/jacob-1.19.jar -DgroupId=com.jacob -DartifactId=jacob -Dversion=1.19 -Dpackaging=jar)
或者修改pom文件也可以使用，最后需要将解压出来的dll文件(64位操作系统放x64,32位操作系统放x86)放入到jdk文件夹下的bin目录中

打印路径详见PrintController中

此项目为后端代码，使用springboot快速搭建，提供接口使用，无数据库操作

前端代码详见：https://github.com/YunZezzZ/print_web.git
