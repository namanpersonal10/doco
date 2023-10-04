#importing modules 
from datetime import datetime
from time import sleep
from flask import Flask, request, send_from_directory, redirect, render_template, url_for, flash, session
from flask.sessions import SecureCookieSessionInterface
import requests
from sqlalchemy import true
from twilio.twiml.messaging_response import Message,MessagingResponse
from twilio.rest import Client 
import requests
import shutil,os
import aspose.words as aw
import fitz
from PIL import Image
from PIL import ImageFilter
from PIL import ImageColor
import pymongo
import PyPDF2
import docx
import subprocess
import secrets
import string

#DB connection with mongoDB
db_client=pymongo.MongoClient("mongodb://localhost:27017") #MongoDB Connection String
# db=db_client['DoCo'] #Cluster name in the DataBase 
# customer_data=db['customer-data'] #Collection Name in the Cluster, importing collection named customer-data

#will update later, till now just you need to know these are used to download the media sent by the user
headers = {
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Safari/537.36'
}

#Function to directly repond any text to user who messaged us
def respond(message):
    response = MessagingResponse()
    response.message(message)
    return str(response)

#account API tokens for Twilio WhatsApp API
account_sid = 'AC33d7286c972802182b4e70140c2fb9a3' 
auth_token = 'bc6ad87b9865f1f49221a1fec09f98c3'

#Twilio Client defined
client = Client(account_sid, auth_token) 

#Flask app defined
app = Flask(__name__)

#pages for error handling
@app.errorhandler(404)
def page_not_found(e):
    #handle 404 error by showing custom 404 error page
    return render_template('404.html'), 404

@app.errorhandler(500)
def server_error(e):
    #handle 500 error by showing custom 500 error page
    return render_template('500.html'), 500

@app.errorhandler(403)
def forbidden(e):
    #handle 403 error by showing custom 403 error page
    return render_template('403.html'), 403

#website home page
@app.route('/')
def home():
    #render home page
    return render_template("index.html")

#website privacy policy page
@app.route('/PrivacyPolicy')
def pp():
    #render privacy policy page
    return render_template("privacy-policy.html")

#Blog Lists Page
@app.route('/Blog')
def Blog():
    #get all blogs from the database and render the blog page
    # blog_list=list(db['BlogList'].find())
    return render_template("blog.html")

#redirecting if someone used /blog instead of /Blog
@app.route('/blog')
def blog():
    #redirect /blog to /Blog
    return redirect('/Blog',302)

#website services page
@app.route('/Services')
def services():
    #adding services, 1. Service Name, 2. Files Supported, 3. Image File name 
    services=[["Word", ["PDF"],"word"],
    ["PDF",["Word","Image","Excel","PowerPoint"],"pdf"],
    ["Image",["PDF","Word","Excel","PowerPoint"],"image"],
    ["Excel",["PDF"],"excel"],
    ["PowerPoint Presentation",["PDF"],"ppt"],
    ["Compressed Doc",["PDF","Image","Excel","PPT"],"compress"]]
    return render_template("service.html", services=services)

#for ads
@app.route('/d755ad6ff4cbdf04f092da39bae9c743.html')
def d755ad6ff4cbdf04f092da39bae9c743():
    return render_template("d755ad6ff4cbdf04f092da39bae9c743.html")

#website single dynamic blog page 
@app.route('/blog/<blog_name>')
def full_blog(blog_name):
    #gathering blog html file name from DB collection named BlogList
    # blog_list=db['BlogList'].find_one({"blog_name":blog_name})
    # if blog_list:
    #     #sending blog page link which is already made
    #     return render_template("blog/"+blog_list['html_file_name'])

    #if no such blog found redirecting to blogs list
    return redirect("/blog")

#website Terms of Service Page
@app.route('/TermsOfService')
def tos():
    return render_template("tos.html")

#website Social Media Page
@app.route('/social')
def social():   
    return render_template("social.html")

#website Feedback Page
@app.route('/feedback')
def feedback():   
    return render_template("feedback.html")

#word convert Page
@app.route('/c/<f_type>', methods=["GET","POST"])
def web_convert(f_type):
    if request.method == 'POST' or request.method=='GET':
        return render_template("download.html",doc_id="trial.png")
    return str(f_type) + "OK"

#Contact form route to add contact form details to the database
@app.route('/contact_form', methods=["GET","POST"])
def contact():
    # contact_db=db['website_contact_form']
    if request.method=="POST":
        name = request.form['name']
        email = request.form['email']
        subject = request.form['subject']
        message = request.form['message']

        c_insert={
            'name':name,
            'email':email,
            'subject':subject,
            'message':message,
        }
        # contact_db.insert_one(c_insert)
        
        return "OK"
    return "Server Busy! Send a mail to support@do-co.info"

#Feedback form route to add contact form details to the database
@app.route('/feedback_form', methods=["GET","POST"])
def feedback_form():
    # contact_db=db['website_feedback_form']
    if request.method=="POST":
        name = request.form['name']
        email = request.form['email']
        message = request.form['message']

        c_insert={
            'name':name,
            'email':email,
            'message':message,
        }
        # contact_db.insert_one(c_insert)
        
        return "OK"
    return "Server Busy! Send a mail to support@do-co.info"

#welcome message template
def welcome_msg(name=""):
    return f"""Hi {name}! Welcome to DoCo-The first WhatsApp Based Document Converter

Please send us a media file to use the service. Currently Acceptable formats are: PDF, DOCX, JPG, XLS, PPT. 

For Any Feedback Please press the button below or just type "Feedback".
For connecting to us on Social Media press "Connect to Us" button below 
*All the original and converted files will be deleted at our end once it is converted and sent back to you.*"""

#dynamic page to return any file(filename) from folder(folder)
@app.route("/<folder>/<filename>")
def hello_world(folder,filename):
    #returing file from directory, it used for PDF to Image
    return send_from_directory(os.path.join(os.getcwd(),f'static/converted/{folder}'),filename)

#upload page yet to developed
@app.route("/upload/<unique_id>")
def upload(unique_id):
    return render_template("upload.html")

#upload page yet to developed
@app.route("/u/<service>", methods=["GET","POST"])
def web_upload(service):
    return render_template("web_upload.html",service=service)

#upload page yet to developed
@app.route("/download/<doc_id>")
def download(doc_id):
    return render_template("download.html",doc_id=doc_id)


#dynamic page to return any file(filename)
@app.route("/<filename>")
def hello_world2(filename):
    #returning filename, used for all other than the PDF to Image
    return send_from_directory(os.path.join(os.getcwd(),'static/converted'),filename)

#dynamic page to redirect to any backlink
@app.route("/l/<back_link>")
def hello_world3(back_link):
    #checking for the forwarding link in the DB collection named link_forwarding
    # user_file_link_collection=db["link_forwarding"] 
    # forwarding_link=user_file_link_collection.find_one({'back_link':str(back_link)})['forward_link']
    #redirecting to fetched link
    # return redirect(forwarding_link)
    return "OK"

#main bot page, here post request is sent when any user sends any message to doco whatsapp number
@app.route("/bot", methods=["GET","POST"]) 
def bot():
    return respond("Currently we are out of service! we will come back shortly!")
    #all requests coming from whatsapp are POST requests
    #if anyone directly tries to access this page then it will redirect to home page of our website
    if request.method=="GET":
        return redirect("/")
    
    #getting values from the JSON object of whatsapp message
    user_msg = request.values.get('Body', '').lower() #user msg, always string, in case of text direct user msg, in case of file it is the filename
    user_num=str(request.values.get('From', '')) #user whatsapp number
    user_name=request.values.get('ProfileName') #user profile name

    if user_num[-10:]=="8054041010":
        try:
            message = client.messages.create(
                        from_='whatsapp:+14155238886', 
                        body = f"hi{user_msg}",    
                        to=f'whatsapp:+918054041010'
                    )
            message = client.messages.create(
                        from_='whatsapp:+14155238886', 
                        body = welcome_msg("Buddy"),    
                        to=f'whatsapp:+91{user_msg}'
                    )
        except Exception as e:
              message = client.messages.create(
                        from_='whatsapp:+14155238886', 
                        body = str(e),    
                        to=user_num
                    )
    if user_num[-10:]=="9878933996":
        return 
    c_user={"name":user_name.lower(),
            "user_num":user_num}
    print(c_user)
    check=customer_data.find_one({"user_num":user_num}) #checking whether customer is new or old
    
    if check==None:
        customer_data.insert_one(c_user) #for new user adding data in the collection

    cus_data = customer_data.find_one({"user_num":user_num})
    
    t_insert={"f_name":"temp",
            "f_type":"temp"}

    user_file_collection=db[str(cus_data['_id'])+"_files"] #opening/creating files collection for the particular user
    all_file=user_file_collection.find()
    for item in all_file:
        if item['f_name']=="temp":
            break
        else: 
            user_file_collection.insert_one(t_insert)
            break

    t_insert_link={"f_name":"temp",
            "f_type":"temp",
            "f_link":"temp"}
    user_file_link_collection=db[str(cus_data['_id'])+"_files_links"] #opening/creating files link collection for the particular user
    all_file_1=user_file_link_collection.find()
    for item in all_file_1:
        if item['f_name']=="temp":
            break
        else: 
            user_file_link_collection.insert_one(t_insert_link)
            break

    #checking whether previously any file is sent by user or not, if yes then getting last file name
    all_file_1=list(user_file_link_collection.find())
    if len(all_file_1)==0:
        user_file_link_collection.insert_one(t_insert_link)
        last_file_name_link=="temp"
    else:
        last_file_name_link = all_file_1[-1]['f_name']
        last_file_name_id = all_file_1[-1]['_id']
        last_file_name_f_link = all_file_1[-1]['f_link']

    if last_file_name_link=="temp":
        file_type=None
    
    user_file_collection=db[str(cus_data['_id'])+"_files"]
   

    user_collection=db[str(cus_data['_id'])]
    f_msg={"msg_body":user_msg,
            "dt":datetime.now()}
    user_collection.insert_one(f_msg)

    #getting number of files sent by a user in a whatsapp message
    num_media = int(request.values.get("NumMedia"))


    all_msg=list(user_collection.find())
    
    all_file=user_file_collection.find()

    last_file_name=""
    for item in all_file:
        if item['f_name']!="temp":
            last_file_name = item['f_name']
            break 
    if last_file_name=="":
        file_type=None
    else:
        if last_file_name[-4:]==".pdf":
            file_type="pdf"
        elif last_file_name[-5:]==".docx":
            file_type="word"
        elif last_file_name[-5:]==".jpeg":
            file_type="image"
        else:
            file_type="unsupported"
    
    #if user only sent a text msg
    if num_media==0:
        #respoding differently upon different type of msgs sent by the user
        if user_msg=="hy" or user_msg=="hello" or user_msg=="hi":
            return respond(welcome_msg(request.values.get('ProfileName')))
        
        elif user_msg=="feedback":
            return respond('Please type your feedback.')


        elif request.values.get("ButtonText"):
            bt=request.values.get("ButtonText").lower()
            if bt=="feedback":
                return respond('Please type your feedback.')
            elif bt=="connect to us":
                return respond('Find us on instagram.\nhttps://www.instagram.com/doco_wa/')
            
            elif bt=="file link" and file_type==None:
                return respond("""Please send a file first or type "hi" to know more""")
            
            elif bt=="file link" and file_type!=None and file_type!="image":
                
                forward_links=db['link_forwarding']
                last_file_name_id_temp=str(last_file_name_id)
                to_insert={'back_link':last_file_name_id_temp,
                            'forward_link':last_file_name_f_link}
                forward_links.insert_one(to_insert)
                
                message = client.messages.create(
                    from_='whatsapp:+14155238886', 
                    body = f"https://do-co.info/l/{last_file_name_id_temp}",    
                    to=user_num
                )
                user_file_link_collection.delete_one({'_id':last_file_name_id})
                return respond("""Please find the above link to share your file.""")
            

            elif bt=="word" and file_type=='pdf':
                user_file_link_collection.delete_one({'f_name':last_file_name})

                respond("Please wait your file is being converted.")
                PDF_to_DOCX(last_file_name,user_num=user_num)   
             

                path = os.path.join(os.getcwd(),'static/upload')
                file_path = os.path.join(path,item['f_name'])
                os.remove(file_path)
                sleep(2)
                path = os.path.join(os.getcwd(),'static/converted')
                file_path = os.path.join(path,item['f_name'][:-4]+".docx")
                os.remove(file_path)
                
                user_file_collection.delete_one({'f_name':item['f_name']})
                
                return respond("Here are your converted Word Document")
            
            elif bt=="word" and file_type==None:
                return respond("""Please send a file first or type "hi" to know more""")

            elif bt=='image' and file_type=="pdf":
                user_file_link_collection.delete_one({'f_name':last_file_name})
                respond("Please wait your file is being converted.")
                PDF_to_IMG(last_file_name,user_num=user_num)

                path = os.path.join(os.getcwd(),'static/upload')
                file_path = os.path.join(path,item['f_name'])
                os.remove(file_path)
                sleep(2)
                path = os.path.join(os.getcwd(),'static/converted')
                file_path = os.path.join(path,item['f_name'][:-4])
                shutil.rmtree(file_path)
                
                user_file_collection.delete_one({'f_name':item['f_name']})
                
                return respond("Here are your converted Images")
            
            elif bt=="image" and file_type==None:
                return respond("""Please send a file first or type "hi" to know more""")
            
            elif bt=='pdf' and file_type=='word':
                user_file_link_collection.delete_one({'f_name':last_file_name})
                message = client.messages.create(
                    from_='whatsapp:+14155238886', 
                    body = "Please wait file is being converted.",    
                    to=user_num
                )
                # respond("Please wait your file is being converted.")
                DOCX_to_PDF(last_file_name,user_num=user_num)

                path = os.path.join(os.getcwd(),'static/upload')
                file_path = os.path.join(path,item['f_name'])
                os.remove(file_path)
                sleep(2)
                

                path = os.path.join(os.getcwd(),'static/converted')
                file_path = os.path.join(path,item['f_name'][:-5]+".pdf")
                os.remove(file_path)
                
                
                user_file_collection.delete_one({'f_name':item['f_name']})

                
                return respond("Here is your converted PDF")

            elif bt=="pdf" and file_type==None:
                return respond("""Please send a file first or type "hi" to know more""")

            elif bt=='image' and file_type=='word':
                user_file_link_collection.delete_one({'f_name':last_file_name})
                respond("Please wait your file is being converted.")
                DOCX_to_IMG(last_file_name,user_num=user_num)   
                

                path = os.path.join(os.getcwd(),'static/upload')
                file_path = os.path.join(path,item['f_name'])
                os.remove(file_path)
                path = os.path.join(os.getcwd(),'static/upload')
                file_path = os.path.join(path,item['f_name'][:-5]+".pdf")
                os.remove(file_path)
                sleep(2)
                path = os.path.join(os.getcwd(),'static/converted')
                file_path = os.path.join(path,item['f_name'][:-4])
                shutil.rmtree(file_path)
                
                user_file_collection.delete_one({'f_name':item['f_name']})
                
                return respond("Here is your converted Images")

            elif bt=='image' and file_type==None:
                return respond("""Please send a file first or type "hi" to know more""")
            
            elif bt=='backspace' and file_type=="image":
                all_file_count=user_file_collection.count_documents({'f_type':'image'})
                path = os.path.join(os.getcwd(),'static/upload')
                file_path = os.path.join(path,str(cus_data['_id'])+f"_{all_file_count-1}.jpeg")
                os.remove(file_path)
               
                user_file_collection.delete_one({'f_name':str(cus_data['_id'])+f"_{all_file_count-1}.jpeg"})
                
                return respond("""Last Image Deleted.\nPress done to convert all sent images to PDF.\n*Press Clear or type 'clear' to clear all images*\nPress Backspace or type 'backspace' to discard last sent image. """)

            elif bt=='clear' and file_type=="image":
                all_file_count=user_file_collection.count_documents({'f_type':'image'})

                path = os.path.join(os.getcwd(),'static/upload')
                for i in range(all_file_count):
                    file_path = os.path.join(path,str(cus_data['_id'])+f"_{i}.jpeg")
                    os.remove(file_path)

                    user_file_collection.delete_one({'f_name':str(cus_data['_id'])+f"_{i}.jpeg"})
                
                return respond("""All Images Deleted. Send File to use service again.\nType *hi* to know about us.""")
            
            elif bt=='done' and file_type=="image":
                respond("Please wait your file is being converted.")
                all_files=user_file_collection.find({'f_type':'image'},{'f_name':1})

                file_list=[]
                for file in all_files:
                    file_list.append(file["f_name"])
    
                IMG_to_PDF(file_list,user_num)

                path = os.path.join(os.getcwd(),'static/upload')
                for i in range(all_file_count):
                    file_path = os.path.join(path,str(cus_data['_id'])+f"_{i}.jpeg")
                    os.remove(file_path)

                    user_file_collection.delete_one({'f_name':str(cus_data['_id'])+f"_{i}.jpeg"})
                
                sleep(2)
                path = os.path.join(os.getcwd(),'static/converted')
                file_path = os.path.join(path,str(cus_data['_id'])+".pdf")
                os.remove(file_path)
                
                return respond("""Here is your converted PDF File.""")


            elif (bt=='clear' or bt=='backspace' or bt=="done") and file_type==None:
                return respond("""Please send a file first or type "hi" to know more""")
            
            
        elif user_msg=="connect to us":
            return respond('Find us on instagram.\nhttps://www.instagram.com/doco_wa/')
        
        elif user_msg=="file link" and file_type!=None and file_type!="image":
                
            forward_links=db['link_forwarding']
            last_file_name_id_temp=str(last_file_name_id)
            to_insert={'back_link':last_file_name_id_temp,
                        'forward_link':last_file_name_f_link}
            forward_links.insert_one(to_insert)

            
            message = client.messages.create(
                from_='whatsapp:+14155238886', 
                body = f"https://do-co.info/l/{last_file_name_id_temp}",    
                to=user_num
            )
            user_file_link_collection.delete_one({'_id':last_file_name_id})
            user_file_collection.delete_one({'f_name':last_file_name})
            return respond("""Please find the above link to share your file.""")
         
        elif (user_msg=='1' or user_msg=='word') and file_type=='pdf':
            user_file_link_collection.delete_one({'f_name':last_file_name})
            respond("Please wait your file is being converted.")
            PDF_to_DOCX(last_file_name,user_num=user_num)   
             

            path = os.path.join(os.getcwd(),'static/upload')
            file_path = os.path.join(path,item['f_name'])
            os.remove(file_path)
            sleep(2)
            path = os.path.join(os.getcwd(),'static/converted')
            file_path = os.path.join(path,item['f_name'][:-4]+".docx")
            os.remove(file_path)
            
            user_file_collection.delete_one({'f_name':item['f_name']})
            
            return respond("Here is your converted Word Document")

        elif (user_msg=='1' or user_msg=='word') and file_type==None:
            return respond("""Please send a file first or type "hi" to know more""")
        
        elif (user_msg=='2' or user_msg=='image') and file_type==None:
            return respond("""Please send a file first or type *hi* to know more""")


        elif (user_msg=='2' or user_msg=='image') and file_type=="pdf":
            user_file_link_collection.delete_one({'f_name':last_file_name})
            respond("Please wait your file is being converted.")
            PDF_to_IMG(last_file_name,user_num=user_num)

            path = os.path.join(os.getcwd(),'static/upload')
            file_path = os.path.join(path,item['f_name'])
            os.remove(file_path)
            sleep(2)
            path = os.path.join(os.getcwd(),'static/converted')
            file_path = os.path.join(path,item['f_name'][:-4])
            shutil.rmtree(file_path)
            
            user_file_collection.delete_one({'f_name':item['f_name']})
            
            return respond("Here are your converted Images")
        
        elif user_msg=='3' or user_msg=='4':
            user_file_link_collection.delete_one({'f_name':last_file_name})
            return respond("""Sorry Currently we're facing issues with PDF to PPT and PDF to Excel, inconvenience is regretful.""")
        
        elif (user_msg=='1' or user_msg=='pdf') and file_type=='word':
            user_file_link_collection.delete_one({'f_name':last_file_name})
            respond("Please wait your file is being converted.")
            DOCX_to_PDF(last_file_name,user_num=user_num)   
             

            path = os.path.join(os.getcwd(),'static/upload')
            file_path = os.path.join(path,item['f_name'])
            os.remove(file_path)
            sleep(2)
            path = os.path.join(os.getcwd(),'static/converted')
            file_path = os.path.join(path,item['f_name'][:-5]+".pdf")
            os.remove(file_path)
            
            user_file_collection.delete_one({'f_name':item['f_name']})
            
            return respond("Here is your converted PDF")

        elif (user_msg=='1' or user_msg=='pdf') and file_type==None:
            return respond("""Please send a file first or type "hi" to know more""")
        
        elif (user_msg=='2' or user_msg=='image') and file_type=='word':
            user_file_link_collection.delete_one({'f_name':last_file_name})
            respond("Please wait your file is being converted.")
            DOCX_to_IMG(last_file_name,user_num=user_num)   
             

            path = os.path.join(os.getcwd(),'static/upload')
            file_path = os.path.join(path,item['f_name'])
            os.remove(file_path)
            path = os.path.join(os.getcwd(),'static/upload')
            file_path = os.path.join(path,item['f_name'][:-5]+".pdf")
            os.remove(file_path)
            sleep(2)
            path = os.path.join(os.getcwd(),'static/converted')
            file_path = os.path.join(path,item['f_name'][:-4])
            shutil.rmtree(file_path)
            
            user_file_collection.delete_one({'f_name':item['f_name']})
            
            return respond("Here is your converted Images")

        elif (user_msg=='2' or user_msg=='image') and file_type==None:
            return respond("""Please send a file first or type "hi" to know more""")
        
        
        elif user_msg=='backspace' and file_type=="image":
            all_file_count=user_file_collection.count_documents({'f_type':'image'})
            path = os.path.join(os.getcwd(),'static/upload')
            file_path = os.path.join(path,str(cus_data['_id'])+f"_{all_file_count-1}.jpeg")
            os.remove(file_path)
            
            user_file_collection.delete_one({'f_name':str(cus_data['_id'])+f"_{all_file_count-1}.jpeg"})
            
            return respond("""Last Image Deleted.\nPress done to convert all sent images to PDF.\n*Press Clear or type 'clear' to clear all images*\nPress Backspace or type 'backspace' to discard last sent image.""")

        elif user_msg=='clear' and file_type=="image":
            all_file_count=user_file_collection.count_documents({'f_type':'image'})

            path = os.path.join(os.getcwd(),'static/upload')
            for i in range(all_file_count):
                file_path = os.path.join(path,str(cus_data['_id'])+f"_{i}.jpeg")
                os.remove(file_path)

                user_file_collection.delete_one({'f_name':str(cus_data['_id'])+f"_{i}.jpeg"})
            
            return respond("""All Images Deleted. Send File to use service again.\nType *hi* to know about us.""")
        
        elif user_msg=='done' and file_type=="image":
            respond("Please wait your file is being converted.")
            all_file_count=user_file_collection.count_documents({'f_type':'image'})

            IMG_to_PDF(str(cus_data['_id']),all_file_count,user_num)

            path = os.path.join(os.getcwd(),'static/upload')
            for i in range(all_file_count):
                file_path = os.path.join(path,str(cus_data['_id'])+f"_{i}.jpeg")
                os.remove(file_path)

                user_file_collection.delete_one({'f_name':str(cus_data['_id'])+f"_{i}.jpeg"})
            sleep(2)
            path = os.path.join(os.getcwd(),'static/converted')
            file_path = os.path.join(path,str(cus_data['_id'])+".pdf")
            os.remove(file_path)
            
            return respond("""Here is your converted PDF File.""")


        elif (user_msg=='clear' or user_msg=='backspace' or user_msg=="done") and file_type==None:
            return respond("""Please send a file first or type "hi" to know more""")
        
        
            
        elif all_msg[-2]['msg_body']=="feedback":
            feed={"feedback":user_msg,
                    "user_num":user_num}
            feedback_collection=db['feedback']
            feedback_collection.insert_one(feed)
            return respond("Thanks for your Feedback, we will surely work on it.")
        else:

            return respond("""Please Send *Hi* or send a file to use the service.""")
    #if user sends a file
    else:
        #getting file type and file url
        media_type=request.values.get("MediaContentType0")
        media_url=request.form.get('MediaUrl0')
        
        # Hy {{1}}, if you wish to make sharable link press the button below.
        message = client.messages.create(
                from_='whatsapp:+14155238886', 
                body = f"Hy {user_name}, if you wish to make sharable link press the button below.",    
                to=user_num
            )
        
        all_file=user_file_collection.find()

        #checking file type and giving files a name
        if str(media_type)=="image/jpeg":

            for item in all_file:
                if item['f_name']!="temp" and item['f_type']!="image":
                    path = os.path.join(os.getcwd(),'static/upload')
                    file_path = os.path.join(path,item['f_name'])
                    try:
                        os.remove(file_path)
                    except:
                        user_file_collection.delete_one({'f_name':item['f_name']})
                        user_file_link_collection.delete_one({'f_name':item['f_name']})
                    user_file_collection.delete_one({'f_name':item['f_name']})

            all_file=list(user_file_collection.find())
        else:
            t_insert_link={"f_name":user_msg,
                "f_type":"known",
                "f_link":media_url}
            user_file_link_collection.insert_one(t_insert_link)
            for item in all_file:
                if item['f_name']!="temp":
                    path = os.path.join(os.getcwd(),'static/upload')
                    file_path = os.path.join(path,item['f_name'])
                    try:
                        os.remove(file_path)
                    except:
                        pass
                    user_file_collection.delete_one({'f_name':item['f_name']})
                    user_file_link_collection.delete_one({'f_name':item['f_name']})

            all_file=list(user_file_collection.find())

        user_msg="_".join( user_msg.split() )
       
        #list of accepted docs
        accepted_docs=["application/vnd.openxmlformats-officedocument.wordprocessingml.document","application/pdf","application/vnd.openxmlformats-officedocument.presentationml.presentation","image/jpeg","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"]
        print(media_type)
        #checking if accepted doc is sent or not, if not then will directly send an option to generate link
        if str(media_type) in accepted_docs:
            if str(media_type)=="application/vnd.openxmlformats-officedocument.presentationml.presentation":
                return respond("We're currently facing issues with PPT and Excel Docs, please try again later.")
            elif str(media_type)== "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
                return respond("hi")
            elif str(media_type)=="image/jpeg":
                t_insert={"f_name":"",
                "f_type":"image"}
                _id=user_file_collection.insert_one(t_insert).inserted_id
                user_msg=str(cus_data['_id'])+"_"+str(_id)+".jpeg"
                

            #saving file to local from the url
            path = os.path.join(os.getcwd(),'static/upload')
            attempts = 0
            while attempts < 5:
                try:
                    filename = user_msg
                    r = requests.get(media_url,headers=headers,stream=True,timeout=5)
                    if r.status_code==200:
                        open(os.path.join(path,filename), 'wb').write(r.content)
                        # with open(os.path.join(path,filename),'wb') as f:
                        #     print('here 1')
                        #     r.raw.decode_content = True
                        #     shutil.copyfileobj(r.raw,f)
                        break
                except Exception as e:
                    print(e)
                    attempts+=1
            #if 5 attempts fot failed will revert for try again
            if attempts==5:
                return respond("Please try to send file again!")
            else:
                while(os.path.exists(os.path.join(path,filename))==0):
                    sleep(2)

            #inserting file name in the DB for further processing
            if str(media_type)=="application/vnd.openxmlformats-officedocument.wordprocessingml.document":
                t_insert={"f_name":user_msg,
                "f_type":"word"}
                user_file_collection.insert_one(t_insert)
                
                return respond(f"Hi {user_name}, we received your Word file select the format you want it to get converted or type number as of below,\n1. PDF\n2. Image")
            elif str(media_type)=="application/pdf":
                t_insert={"f_name":user_msg,
                "f_type":"PDF"}
                user_file_collection.insert_one(t_insert)
                
                return respond(f"Hi {user_name}, we received your PDF file select the format you want it to get converted or type number as of below,\n1. Word\n2. Image\n3. Power Point Presentation\n4. Excel")
            elif str(media_type)=="image/jpeg":  
                
                user_file_collection.update_one({"_id":_id},{ "$set": { 'f_name': user_msg } })
                return respond(f"Hi {user_name}, we received your Image file, Send more images or press done to convert all sent images to PDF.\n*Press Clear or type 'clear' to clear all images*\nPress Backspace or type 'backspace' to discard last sent image.")
        
        else:
            message = client.messages.create(
                from_='whatsapp:+14155238886', 
                body = "Please send an accepted file format.",    
                to=user_num
            )
            return respond(welcome_msg(request.values.get('ProfileName')))
        
        return respond("Please Send *Hi* or send a file to use the service.")

    return respond("Please Send *Hi* or send a file to use the service.")

#function to convert PDF to Word
def PDF_to_DOCX(filename,user_num):
    
    path = os.path.join(os.getcwd(),'static/upload')
    file_path = os.path.join(path,filename)

    path2 = os.path.join(os.getcwd(),'static/converted')
    docx_file_path = os.path.join(path2,filename[:-4]+".docx")

    mydoc = docx.Document()
    pdfFileObj = open(file_path, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

    for pageNum in range(1, pdfReader.numPages):
        pageObj = pdfReader.getPage(pageNum)
        pdfContent = pageObj.extractText()

        mydoc.add_paragraph(pdfContent)

    mydoc.save(docx_file_path)

    sleep(2)
        
    message = client.messages.create(
                from_='whatsapp:+919872110102', 
                media_url = f"https://do-co.info/{filename[:-4]}.docx",    
                to=user_num
    )

    sleep(1)

#funtion to convert Word to PDF
def DOCX_to_PDF(filename,user_num):

    if user_num !='whatsapp:+919717290345':
        message = client.messages.create(
                    from_='whatsapp:+14155238886', 
                    body = "Sorry we are currently facing issues with DOcx to PDF, please try again later.",    
                    to=user_num
                )    
        return 
    path = os.path.join(os.getcwd(),'static/upload')
    file_path = os.path.join(path,filename)

    message = client.messages.create(
                from_='whatsapp:+14155238886', 
                body = "here 2",    
                to=user_num
            )    

    path2 = os.path.join(os.getcwd(),r'static\converted')
    pdf_file_path = os.path.join(path2,filename[:-5]+".pdf")

    message = client.messages.create(
                from_='whatsapp:+14155238886', 
                body = "here 3",    
                to=user_num
            )    
        
    message = client.messages.create(
            from_='whatsapp:+14155238886', 
            body = "here 4",    
            to=user_num
        )    
    
    os.system(f'libreoffice --convert-to pdf /home/naman/doco/static/upload/{filename} --outdir /home/naman/doco/static/converted')

    message = client.messages.create(
            from_='whatsapp:+14155238886', 
            body = "here 111",    
            to=user_num
        )

    sleep(2)
    
    message = client.messages.create(
                from_='whatsapp:+14155238886', 
                body = "here 7",    
                to=user_num
            )    
    
    message = client.messages.create(
                from_='whatsapp:+14155238886', 
                media_url = f"https://do-co.info/{filename[:-5]}.pdf",    
                to=user_num
    )

    
    message = client.messages.create(
                from_='whatsapp:+14155238886', 
                body = "here 8",    
                to=user_num
            )    

    sleep(1)

#funtion to convert Image to PDF
def IMG_to_PDF(file_list,user_num):
    
    path = os.path.join(os.getcwd(),'static/upload')
    first_img_path = os.path.join(path,file_list[0])
    first_img=Image.open(first_img_path)
    img_1=first_img.convert('RGB')

    img_list=[]

    for i in range(1,len(file_list)):
        img_path = os.path.join(path,file_list[i])
        img_open=Image.open(img_path)
        img_convert=img_open.convert('RGB')

        img_list.append(img_convert)
    
    path2 = os.path.join(os.getcwd(),'static/converted')
    pdf_file_path= os.path.join(path2,file_list[0]+".pdf")
    img_1.save(pdf_file_path,save_all=True,append_images=img_list)    
    
    sleep(2)
    
    message = client.messages.create(
                from_='whatsapp:+14155238886', 
                media_url = f"https://do-co.info/{file_list[0]}.pdf",    
                to=user_num
    )

    sleep(1)

#funtion to convert Word to Image
def DOCX_to_IMG(filename,user_num):
    
    path = os.path.join(os.getcwd(),'static/upload')
    file_path = os.path.join(path,filename)


    path2 = os.path.join(os.getcwd(),r'static\upload')
    pdf_file_path = os.path.join(path2,filename[:-5]+".pdf")
    
    os.system(f'libreoffice --headless --convert-to pdf {file_path} --outdir /home/naman/doco/static/upload')

    PDF_to_IMG(filename=filename[:-5]+".pdf",user_num=user_num)
    sleep(1)

#funtion to convert PDF to Image
def PDF_to_IMG(filename,user_num):
    
    path = os.path.join(os.getcwd(),'static/upload')
    file_path=os.path.join(path,filename)

    
    pdf_file = fitz.open(file_path)
    
    total_page = len(pdf_file)
    
    os.mkdir(os.path.join(os.getcwd(),f'static/converted/{filename[:-4]}'))
    
    mat = fitz.Matrix(2.0, 2.0)
    
    for page_index in range(total_page):
    
        cpage = pdf_file[page_index]

        cpage_img = cpage.get_pixmap(matrix=mat)

        img = Image.frombytes("RGB", [cpage_img.width, cpage_img.height], cpage_img.samples)

        cimg_name = f"{filename[:-4]}_{str(page_index)}.png"
        img.save(os.path.join(os.getcwd(),f'static/converted/{filename[:-4]}/{cimg_name}'))
        
        message = client.messages.create(
                    from_='whatsapp:+14155238886',  
                    body=f"Page - {page_index+1}",  
                    media_url = "https://do-co.info/"+filename[:-4]+"/"+cimg_name,    
                    to=user_num
        )
        sleep(1)

if __name__ == "__main__":
    #running flask app
    app.run(debug=True,host='0.0.0.0')