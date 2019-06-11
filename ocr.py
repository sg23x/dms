from PIL import Image
import pytesseract,openpyxl
from selenium import webdriver
from selenium.webdriver.support.ui import Select
import time,datetime
book=openpyxl.load_workbook("btechlist.xlsx")
sheet=book["Student List"]
newbook=openpyxl.load_workbook("booki.xlsx")
newsheet=newbook["Sheet1"]
print("START")
print(datetime.datetime.now())
x=0
a=webdriver.Firefox()
a.get("https://dms.jaipur.manipal.edu/loginForm.aspx")
for i in range(2,1144):
    try:
        dob=sheet["K"+str(i)].value
        username=str(sheet["B"+str(i)].value)
        section=str(sheet["D"+str(i)].value)
        password=(dob.strftime("%d%b"+str(dob.year))).upper()
        a.find_element_by_css_selector("#txtUserid").send_keys(username)
        a.find_element_by_css_selector("#txtpassword").send_keys(password)
        d=a.find_element_by_css_selector("#imgCaptcha")
        loc=d.location
        size=d.size
        a.save_screenshot("scree1.png")
        imgg=Image.open("scree1.png")
        left=loc["x"]
        top=loc["y"]
        right=left+size["width"]
        bottom=top+size["height"]
        imgg=imgg.crop((left,top,right,bottom))
        imgg.save("scree1.png")
        cap=pytesseract.image_to_string(Image.open("scree1.png"))
        a.find_element_by_css_selector("#txtCaptcha").send_keys(cap)
        a.find_element_by_css_selector("#hprlnkStduent").click()
        time.sleep(2)
        a.find_element_by_css_selector("#reptrMainManu1_lnkbtn_3").click()
        time.sleep(1)
        a.find_element_by_css_selector("#sub-tabs-list > li:nth-child(4) > a:nth-child(1)").click()
        gpa=a.find_element_by_css_selector("#ContentPlaceHolder1_GrdExamResults2 > tbody > tr:nth-child(2) > td:nth-child(2)")
        name=str(sheet["C"+str(i)].value)
        print("Done "+name+" ("+str(i-1)+")")
        newsheet["A"+str(i)]=username
        newsheet["B"+str(i)]=name
        newsheet["C"+str(i)]=section
        newsheet["D"+str(i)]=gpa.text
        newbook.save("booki.xlsx")
        a.find_element_by_css_selector("#pnlImgBlank > i").click()
        a.find_element_by_css_selector("#lnkSignout").click()
    except:
        try:
            a.find_element_by_css_selector("#btnRefresh").click()
            a.find_element_by_css_selector("#txtUserid").clear()
            a.find_element_by_css_selector("#txtpassword").clear()
            a.find_element_by_css_selector("#txtCaptcha").clear()
            dob=sheet["K"+str(i)].value
            username=str(sheet["B"+str(i)].value)
            section=str(sheet["D"+str(i)].value)
            password=(dob.strftime("%d%b"+str(dob.year))).upper()
            a.find_element_by_css_selector("#txtUserid").send_keys(username)
            a.find_element_by_css_selector("#txtpassword").send_keys(password)
            d=a.find_element_by_css_selector("#imgCaptcha")
            loc=d.location
            size=d.size
            a.save_screenshot("scree1.png")
            imgg=Image.open("scree1.png")
            left=loc["x"]
            top=loc["y"]
            right=left+size["width"]
            bottom=top+size["height"]
            imgg=imgg.crop((left,top,right,bottom))
            imgg.save("scree1.png")
            cap=pytesseract.image_to_string(Image.open("scree1.png"))
            a.find_element_by_css_selector("#txtCaptcha").send_keys(cap)
            a.find_element_by_css_selector("#hprlnkStduent").click()
            time.sleep(2)
            a.find_element_by_css_selector("#reptrMainManu1_lnkbtn_3").click()
            time.sleep(1)
            a.find_element_by_css_selector("#sub-tabs-list > li:nth-child(4) > a:nth-child(1)").click()
            gpa=a.find_element_by_css_selector("#ContentPlaceHolder1_GrdExamResults2 > tbody > tr:nth-child(2) > td:nth-child(2)")
            name=str(sheet["C"+str(i)].value)
            print("Done "+name+" ("+str(i-1)+")")
            newsheet["A"+str(i)]=username
            newsheet["B"+str(i)]=name
            newsheet["C"+str(i)]=section
            newsheet["D"+str(i)]=gpa.text
            newbook.save("booki.xlsx")
            a.find_element_by_css_selector("#pnlImgBlank > i").click()
            a.find_element_by_css_selector("#lnkSignout").click()
        except:
            try:
                a.find_element_by_css_selector("#btnRefresh").click()
                a.find_element_by_css_selector("#txtUserid").clear()
                a.find_element_by_css_selector("#txtpassword").clear()
                a.find_element_by_css_selector("#txtCaptcha").clear()
                dob=sheet["K"+str(i)].value
                username=str(sheet["B"+str(i)].value)
                section=str(sheet["D"+str(i)].value)
                password=(dob.strftime("%d%b"+str(dob.year))).upper()
                a.find_element_by_css_selector("#txtUserid").send_keys(username)
                a.find_element_by_css_selector("#txtpassword").send_keys(password)
                d=a.find_element_by_css_selector("#imgCaptcha")
                loc=d.location
                size=d.size
                a.save_screenshot("scree1.png")
                imgg=Image.open("scree1.png")
                left=loc["x"]
                top=loc["y"]
                right=left+size["width"]
                bottom=top+size["height"]
                imgg=imgg.crop((left,top,right,bottom))
                imgg.save("scree1.png")
                cap=pytesseract.image_to_string(Image.open("scree1.png"))
                a.find_element_by_css_selector("#txtCaptcha").send_keys(cap)
                a.find_element_by_css_selector("#hprlnkStduent").click()
                time.sleep(2)
                a.find_element_by_css_selector("#reptrMainManu1_lnkbtn_3").click()
                time.sleep(1)
                a.find_element_by_css_selector("#sub-tabs-list > li:nth-child(4) > a:nth-child(1)").click()
                gpa=a.find_element_by_css_selector("#ContentPlaceHolder1_GrdExamResults2 > tbody > tr:nth-child(2) > td:nth-child(2)")
                name=str(sheet["C"+str(i)].value)
                print("Done "+name+" ("+str(i-1)+")")
                newsheet["A"+str(i)]=username
                newsheet["B"+str(i)]=name
                newsheet["C"+str(i)]=section
                newsheet["D"+str(i)]=gpa.text
                newbook.save("booki.xlsx")
                a.find_element_by_css_selector("#pnlImgBlank > i").click()
                a.find_element_by_css_selector("#lnkSignout").click()
            except:
                
                try:
                    a.find_element_by_css_selector("#btnRefresh").click()
                    a.find_element_by_css_selector("#txtUserid").clear()
                    a.find_element_by_css_selector("#txtpassword").clear()
                    a.find_element_by_css_selector("#txtCaptcha").clear()
                    dob=sheet["K"+str(i)].value
                    username=str(sheet["B"+str(i)].value)
                    section=str(sheet["D"+str(i)].value)
                    password=(dob.strftime("%d%b"+str(dob.year))).upper()
                    a.find_element_by_css_selector("#txtUserid").send_keys(username)
                    a.find_element_by_css_selector("#txtpassword").send_keys(password)
                    d=a.find_element_by_css_selector("#imgCaptcha")
                    loc=d.location
                    size=d.size
                    a.save_screenshot("scree1.png")
                    imgg=Image.open("scree1.png")
                    left=loc["x"]
                    top=loc["y"]
                    right=left+size["width"]
                    bottom=top+size["height"]
                    imgg=imgg.crop((left,top,right,bottom))
                    imgg.save("scree1.png")
                    cap=pytesseract.image_to_string(Image.open("scree1.png"))
                    a.find_element_by_css_selector("#txtCaptcha").send_keys(cap)
                    a.find_element_by_css_selector("#hprlnkStduent").click()
                    time.sleep(2)
                    a.find_element_by_css_selector("#reptrMainManu1_lnkbtn_3").click()
                    time.sleep(1)
                    a.find_element_by_css_selector("#sub-tabs-list > li:nth-child(4) > a:nth-child(1)").click()
                    gpa=a.find_element_by_css_selector("#ContentPlaceHolder1_GrdExamResults2 > tbody > tr:nth-child(2) > td:nth-child(2)")
                    name=str(sheet["C"+str(i)].value)
                    print("Done "+name+" ("+str(i-1)+")")
                    newsheet["A"+str(i)]=username
                    newsheet["B"+str(i)]=name
                    newsheet["C"+str(i)]=section
                    newsheet["D"+str(i)]=gpa.text
                    newbook.save("booki.xlsx")
                    a.find_element_by_css_selector("#pnlImgBlank > i").click()
                    a.find_element_by_css_selector("#lnkSignout").click()
                except:
                    
                    
                    try:
                        a.find_element_by_css_selector("#btnRefresh").click()
                        a.find_element_by_css_selector("#txtUserid").clear()
                        a.find_element_by_css_selector("#txtpassword").clear()
                        a.find_element_by_css_selector("#txtCaptcha").clear()
                        dob=sheet["K"+str(i)].value
                        username=str(sheet["B"+str(i)].value)
                        section=str(sheet["D"+str(i)].value)
                        password=(dob.strftime("%d%b"+str(dob.year))).upper()
                        a.find_element_by_css_selector("#txtUserid").send_keys(username)
                        a.find_element_by_css_selector("#txtpassword").send_keys(password)
                        d=a.find_element_by_css_selector("#imgCaptcha")
                        loc=d.location
                        size=d.size
                        a.save_screenshot("scree1.png")
                        imgg=Image.open("scree1.png")
                        left=loc["x"]
                        top=loc["y"]
                        right=left+size["width"]
                        bottom=top+size["height"]
                        imgg=imgg.crop((left,top,right,bottom))
                        imgg.save("scree1.png")
                        cap=pytesseract.image_to_string(Image.open("scree1.png"))
                        a.find_element_by_css_selector("#txtCaptcha").send_keys(cap)
                        a.find_element_by_css_selector("#hprlnkStduent").click()
                        time.sleep(2)
                        a.find_element_by_css_selector("#reptrMainManu1_lnkbtn_3").click()
                        time.sleep(1)
                        a.find_element_by_css_selector("#sub-tabs-list > li:nth-child(4) > a:nth-child(1)").click()
                        gpa=a.find_element_by_css_selector("#ContentPlaceHolder1_GrdExamResults2 > tbody > tr:nth-child(2) > td:nth-child(1)")
                        name=str(sheet["C"+str(i)].value)
                        print("Done "+name+" ("+str(i-1)+")")
                        newsheet["A"+str(i)]=username
                        newsheet["B"+str(i)]=name
                        newsheet["C"+str(i)]=section
                        newsheet["D"+str(i)]=gpa.text
                        newbook.save("booki.xlsx")
                        a.find_element_by_css_selector("#pnlImgBlank > i").click()
                        a.find_element_by_css_selector("#lnkSignout").click()
                    except:
                        try:
                            a.find_element_by_css_selector("#btnRefresh").click()
                            a.find_element_by_css_selector("#txtUserid").clear()
                            a.find_element_by_css_selector("#txtCaptcha").clear()
                            a.find_element_by_css_selector("#txtpassword").clear()
                            dob=sheet["K"+str(i)].value
                            username=str(sheet["B"+str(i)].value)
                            section=str(sheet["D"+str(i)].value)
                            password=(dob.strftime("%d%b"+str(dob.year))).upper()
                            a.find_element_by_css_selector("#txtUserid").send_keys(username)
                            a.find_element_by_css_selector("#txtpassword").send_keys(password)
                            d=a.find_element_by_css_selector("#imgCaptcha")
                            loc=d.location
                            size=d.size
                            a.save_screenshot("scree1.png")
                            imgg=Image.open("scree1.png")
                            left=loc["x"]
                            top=loc["y"]
                            right=left+size["width"]
                            bottom=top+size["height"]
                            imgg=imgg.crop((left,top,right,bottom))
                            imgg.save("scree1.png")
                            cap=pytesseract.image_to_string(Image.open("scree1.png"))
                            a.find_element_by_css_selector("#txtCaptcha").send_keys(cap)
                            a.find_element_by_css_selector("#hprlnkStduent").click()
                            time.sleep(2)
                            a.find_element_by_css_selector("#reptrMainManu1_lnkbtn_3").click()
                            time.sleep(1)
                            a.find_element_by_css_selector("#sub-tabs-list > li:nth-child(4) > a:nth-child(1)").click()
                            gpa=a.find_element_by_css_selector("#ContentPlaceHolder1_GrdExamResults2 > tbody > tr:nth-child(2) > td:nth-child(1)")
                            name=str(sheet["C"+str(i)].value)
                            print("Done "+name+" ("+str(i-1)+")")
                            newsheet["A"+str(i)]=username
                            newsheet["B"+str(i)]=name
                            newsheet["C"+str(i)]=section
                            newsheet["D"+str(i)]=gpa.text
                            newbook.save("booki.xlsx")
                            a.find_element_by_css_selector("#pnlImgBlank > i").click()
                            a.find_element_by_css_selector("#lnkSignout").click()
                        except:
                            try:
                                a.find_element_by_css_selector("#btnRefresh").click()
                                a.find_element_by_css_selector("#txtUserid").clear()
                                a.find_element_by_css_selector("#txtpassword").clear()
                                a.find_element_by_css_selector("#txtCaptcha").clear()
                                dob=sheet["K"+str(i)].value
                                username=str(sheet["B"+str(i)].value)
                                section=str(sheet["D"+str(i)].value)
                                password=(dob.strftime("%d%b"+str(dob.year))).upper()
                                a.find_element_by_css_selector("#txtUserid").send_keys(username)
                                a.find_element_by_css_selector("#txtpassword").send_keys(password)
                                d=a.find_element_by_css_selector("#imgCaptcha")
                                loc=d.location
                                size=d.size
                                a.save_screenshot("scree1.png")
                                imgg=Image.open("scree1.png")
                                left=loc["x"]
                                top=loc["y"]
                                right=left+size["width"]
                                bottom=top+size["height"]
                                imgg=imgg.crop((left,top,right,bottom))
                                imgg.save("scree1.png")
                                cap=pytesseract.image_to_string(Image.open("scree1.png"))
                                a.find_element_by_css_selector("#txtCaptcha").send_keys(cap)
                                a.find_element_by_css_selector("#hprlnkStduent").click()
                                time.sleep(2)
                                a.find_element_by_css_selector("#reptrMainManu1_lnkbtn_3").click()
                                time.sleep(1)
                                a.find_element_by_css_selector("#sub-tabs-list > li:nth-child(4) > a:nth-child(1)").click()
                                gpa=a.find_element_by_css_selector("#ContentPlaceHolder1_GrdExamResults2 > tbody > tr:nth-child(2) > td:nth-child(1)")
                                name=str(sheet["C"+str(i)].value)
                                print("Done "+name+" ("+str(i-1)+")")
                                newsheet["A"+str(i)]=username
                                newsheet["B"+str(i)]=name
                                newsheet["C"+str(i)]=section
                                newsheet["D"+str(i)]=gpa.text
                                newbook.save("booki.xlsx")
                                a.find_element_by_css_selector("#pnlImgBlank > i").click()
                                a.find_element_by_css_selector("#lnkSignout").click()
                            except:
                                name=str(sheet["C"+str(i)].value)
                                print("FAILED :"+str(name)+" "+str(i-1))
                                pass
            
print("DONE ")
print(datetime.datetime.now())