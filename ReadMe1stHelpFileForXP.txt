This help file is only valid for windows XP.

Download project files from http://www.geocities.com/asakpke/ws/

1. First of all install Internet Information Services (IIS) that will run your
ASP's code. For installing this go control panel -> Add or Remove
Programs -> Add/Remove Windows Components & select check box before
Internet Information Service (IIS). Click Next/Ok & that is it.

2. Go control panel -> Administrative Tools -> Internet Information
Services. Right click on Default Web Site -> New -> Virtual Directory ...
Then give the Alias name 'mcb'. Then browse the mcb website file i.e
IADProjectMCBWebsite\Mcb Website. Then Next Next Finish. Follow same
procedure to create Virtual Directory of shop's website. i.e mcb for
MCB's website, shop for shop's website

3. Create OBDC Data Source from Control Panel (on System DSN tab) for MCB's
database and shop's database.

Note: use data sourse name as used in ASP's file for database connectivity
i.e dsnMCB for MCB's database, dsnShop for Shop's database. Or use new
data sourse name but change data sourse name in each ASP file that use the
database connectivity.

4. Goto DOS. Move control to folder 'My Dll 6 Encription &
Decription'. then type regsvr32 MyDll6.dll. Click ok.


5. Start/run the Internet Explorer.

6. Type the following address on address bar.
http://Your Computer Name/Your VD Name For MCB/Shop.
i.e default address is following
localhost/mcb/
Or
/localhost/shop/

7. Login go website from right side form. by user name 'aamir' &
password 'b'
Create new refe to. & logout. then new customers & enjoy
