Title: [Google Drive] Rename files
Date: 2018-11-22 20:10
Author: Lulef
Category: Sammlung
Slug: google-drive-rename-files
Status: published
Google Apps Script
------------------
function rename() {\
var fldr=DriveApp.getFolderById("1i7Q-8OtwYPe8xjrGMBtHjau7\_vrryg");\
var d=fldr.getFiles();\
var f,n;\
while (d.hasNext()) {\
f=d.next();\
n=f.getName();\
start=n.slice(0,4).toLowerCase()\
if (start.match("img\_\|img-\|vid\_\|pano")) {\
ext=n.slice(-3).toLowerCase();
myYear=n.slice(4,8);\
myMonth=n.slice(8,10);\
myDay=n.slice(10,12);\
myHour=n.slice(13,15);\
myMinute=n.slice(15,17);\
mySecond=n.slice(17,19);
f.setName(myYear+'-'+myMonth+'-'+myDay+' '+myHour+'-'+myMinute+'-'+mySecond+'.'+ext);\
}\
}\
}
