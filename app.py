# Import required libraries
import streamlit as st
import toml
# Import Python Libraries
import openai
import os
#from IPython.display import clear_output
import pandas as pd
#import json
#import io
import xlsxwriter
import requests

from PIL import Image

import io

from IPython.display import display

import random
# Image Generation
from PIL import Image
from openai import OpenAI
# Azure ADLS connection
from azure.identity import ClientSecretCredential
from azure.storage.blob import BlobServiceClient

###############################################################################################################################
st.markdown(
    """<style>
div[class*="stSelectBox"] > label > div[data-testid="stMarkdownContainer"] > p {font-size: 18px;}
div[class*="stSlider"] > label > div[data-testid="stMarkdownContainer"] > p {font-size: 18px;}
div[class*="stTextArea"] > label > div[data-testid="stMarkdownContainer"] > p {font-size: 18x;}
div[class*="stTextInput"] > label > div[data-testid="stMarkdownContainer"] > p {font-size: 18px;}
div[class*="stCheckBox"] > label > div[data-testid="stMarkdownContainer"] > p {font-size: 18px;}
div[class*="stFileUploader"] > label > div[data-testid="stMarkdownContainer"] > p {font-size: 18px;}
div[class*="stTextInput"] > label > div[data-testid="stMarkdownContainer"] > p {font-size: 18px;}
div[class*="stEmpty"] > label > div[data-testid="stMarkdownContainer"] > p {font-size: 18px;}
.stTabs [data-baseweb="tab-list"] button [data-testid="stMarkdownContainer"] p {font-size:22px;}
.block-container {padding-top: 2rem;}
.reportview-container .sidebar-content {{padding-top: 2rem;}}
    </style>
    """, unsafe_allow_html=True)

def check_password():
    """Returns `True` if the user had the correct password."""
    def password_entered():
        """Checks whether a password entered by the user is correct."""
        if st.session_state["password"] == 'icici123#':
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # don't store password
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # First run, show input for password.
        st.text_input(
            "Password", type="password", on_change=password_entered, key="password"
        )
        return False
    elif not st.session_state["password_correct"]:
        # Password not correct, show input + error.
        st.text_input(
            "Password", type="password", on_change=password_entered, key="password"
        )
        st.error("😕 Password incorrect")
        return False
    else:
        # Password correct.
        return True

if check_password():
    st.title("Content design using ChatGPT")

    ###############################################################################################################################
    # Declare chatbot class to interact with
    class ChatBot:
        def __init__(self, num_char = 500):
            # Set the api key
            secret = '9m6nMs0jNBXmsqaP'
            openai.api_key = 'sk-2EEQc18rmdYRXZoCciExT3BlbkFJuLM6' + secret # official
            self.api_key = openai.api_key
            # 1 token = 4 characters = 0.75 words
            self.max_tokens = round(0.25*num_char)

        def get_response(self, prompt, tone, role, product):
            content = "You’re a " + tone + role + 'and have great expertise writing campaigns' + product #"campign message writer"
            messages = [
                {"role": "system", "content" : content}#"You’re a campign message writer"}
            ]
            # Get response for user prompt
            content = prompt
            messages.append({"role": "user", "content": content})

            # Declare the client for chat completion
            client = OpenAI(api_key = self.api_key)
            # Chat Completion
            response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=messages
            )

            chat_response = response.choices[0].message.content
            return chat_response

    class FreeStyle:
        def __init__(self, num_char = 500):
            # Set the api key
            secret = '9m6nMs0jNBXmsqaP'
            openai.api_key = 'sk-2EEQc18rmdYRXZoCciExT3BlbkFJuLM6' + secret # official
            self.api_key = openai.api_key
            # 1 token = 4 characters = 0.75 words
            self.max_tokens = round(0.25*num_char)
        def get_response(self, prompt):
            content = '''You are a professional ICICI Bank executive trained for writing campaigns content. You will respond to users queries in writing campaign/marketing
                messages/SMS/email/notification. In case of requests that are not related to campaign content design,
                You will politely refuse their request telling them that you have expertise only in writing campaign or
                marketing messages and can not help with other subjects or tasks'''
            messages = [
                {"role": "system", "content" : content}#"You’re a campign message writer"}
            ]
            # Get response for user prompt
            content = prompt
            messages.append({"role": "user", "content": content})

            # Declare the client for chat completion
            client = OpenAI(api_key = self.api_key)
            # Chat Completion
            response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=messages
            )

            chat_response = response.choices[0].message.content
            return chat_response

    def to_excel(df):
        output_filename = "output.xlsx"  # Replace with desired output file name
        writer = pd.ExcelWriter(output_filename, engine='xlsxwriter')
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        format1 = workbook.add_format({'num_format': '0.00'})
        worksheet.set_column('A:A', None, format1)
        writer.save()
        with open(output_filename, 'rb') as file:
            processed_data = file.read()
        return processed_data
    ###############################################################################################################################
    # Tuple of segments
    seg_tuple = ("NONE", "ACADEMIC", "ADVENTURER", "ASPIRATIONAL", "BOHEMIAN", "CORPORATE", "COSMOPOLITAN", "CREATIVE_IMAGINE",
                            "NATURE_LOVER", "PROVIDER", "SOCIALITE", "SOCIALLY_CONCERNED", "TECHNICAL_PROFESSIONALS")
    campaign_list = ['SMS', 'Email', 'Headline', 'Notification']
    product_tuple = ("General", "Agri Corporates", "ATM", "Bill Pay", "Bonds", "Business Loan", "Branch Banking", "Car Loan", "Commercial Business Loan",
                     "Commercial Card", "Credit Card", "Current Account", "Debit Card", "Education Loan",
                     "Farmer Finance", "Fixed Deposit", "Forex Card", "Gold Loan", "Gold Monetisation Scheme",
                     "Goods and Services Tax (GST)", "Health Insurance", "Healthcare Equipment Loan", "HL Topup",
                     "Home Loan", "Home Loan Balance Transfer", "Instant Gold Loan", "Internet Banking",
                     "Investment and Tax Savings Account", "IPO through ASBA (Applications Supported by Blocked Amount)",
                     "iWish", "Life Insurance", "Loans Against Securities", "Merchant Services", "Micro Banking",
                     "Mobile Banking", "Mutual Fund", "National Pension System", "Personal Loan", "PL on Credit Card",
                     "Pockets", "Pradhan Mantri Mudra Yojana (PMMY)", "Prepaid Card", "Public Provident Fund",
                     "Recurring Deposit", "Rural & Agri Business", "Salary Overdraft", "Savings Account", "Senior Citizen Saving Scheme",
                     "SIP", "Stand-Up India Scheme", "Stocks", "Sukanya Samriddhi Yojana (SSY) Account", "Tax e-filing",
                     "Tax Solutions", "Tractor Loan", "Travel Card", "Travel Insurance", "Two Wheeler Loan",
                     "Unifare Metro Card")

    tone_tuple = ('Encouraging', 'Motivational', 'Exciting', 'Formal', 'Informal', 'Friendly', 'Optimistic', 'Assertive')
    char_tuple = tuple([i for i in range(50,550,50)])
    role_tuple = ('Customer Service Agent', 'Campaign Writer', 'Customer Relationship Manager')
    age_tuple = ('NONE', '23 years – 30 years', '30 years – 40 years', '40 years – 50 years', '50 years – 65 years')
    income_tuple = ("NONE", "40k-70k","70k – 1L","1L-1.5L","1.5L-2L","2L-2.5L","2.5L-3L",">3L")
    # Prompt for each segment
    segments = ['', 'in academic tone', 'in adventurous tone',
                'in aspirational tone',
                'in creative tone', 'for corporate employees but do not address them as corporate employees',
                'a cosmopolitan tone but do not address them as comsopolitans', 'in a creative & imaginative tone',
                'for nature loving people but do not address them as nature lovers','for family supporters but do not address as family supporters',
                'for social people but do not address them as social people',
                'for socially concious people but do not address as socially concious',
                'for technical professionals but do not address them as technical proffessionals']
    products = [""]+["for ICICI Bank " + i for i in product_tuple[1:]]
    role_content = ['You are a customer service agent','You are a campaign writer','You are a customer relationship manager']
    age_list = ['', '. They are also Young Adults but do not address as young adults',
                '. They are also Early Adults but do not address as early adults',
                '. They are also Middle Adults but do not address as middle adults',
                '. They are also Late adults or senior but do not address as late adults or senior']
    income_list = ["", ". They have Lower Middle Income",". They have Middle Income",". They have Upper Middle Income",". They have Affluent income",
                   ". They have High Income",". They have Upper High Income",". They have Very High Income"]
    ###############################################################################################################################
    # Create a sidebar
    st.sidebar.header("About the App")
    st.sidebar.info('''Welcome to the ICICI Bank Content Design Application.
    This web application allows you to generate personalized content for ICICI Bank customers using the power of ChatGPT.''')
    st.sidebar.header("Interactive Mode")
    st.sidebar.info('''
        1. Select a psychographic segment, campaign type and role, tone, maximum response characters.
        2. Enter your message in the text box and hit send
        3. Click "Clear to clean up the chat"''')

    st.sidebar.header("Power Mode")

    st.sidebar.info('''
        1. Upload an Excel file
        2. Click "Generate Response"
        3. Download Batch Results''')

    bank_logo = 'https://upload.wikimedia.org/wikipedia/commons/thumb/1/12/ICICI_Bank_Logo.svg/2560px-ICICI_Bank_Logo.svg.png'
    st.sidebar.image(bank_logo, use_column_width=True, width=150)

    st.sidebar.text('Developed by DSAG')

    ###############################################################################################################################
    # CREATE TABS
    tabs = st.tabs(['Interactive Mode', 'Power Mode', 'Image Generation'])
    with tabs[0]:
        # Subheading
        # Free Style CheckBox
        free_style_mode = st.checkbox("Enable Free Style Mode", key="disabled")
        # Drop downs and sliders
        if not free_style_mode:
            x = st.selectbox('Psychographic Segment', seg_tuple)
            campaign_type = st.selectbox('Campaign Type', campaign_list)
            product = st.selectbox('Product', product_tuple)
            role = st.selectbox('Role', role_tuple)
            age = st.selectbox('Age Band', age_tuple)
            income = st.selectbox('Income Band', income_tuple)
            cam_tone = st.selectbox('Campaign Tone', tone_tuple)
            num_char = st.slider('Maximum Response Characters', min_value=30, max_value=10000, step=1)
            num_word = str(int(round(num_char/6, 0)))
            # Use index of role_tuple to get corresponding role_content
            idx2 = role_tuple.index(role)
            content = role_content[idx2]
            # Use index of seg_tuple to get corresponding segment prompt
            idx1 = seg_tuple.index(x)
            seg = segments[idx1]
            # Use index of product_tuple to get corresponding product statements
            idx_prod = product_tuple.index(product)
            product_indexed = products[idx_prod]
            # Use index of age_tuple to get corresponding age prompts
            idx_age = age_tuple.index(age)
            age_indexed = products[idx_age]
            # Use index of income_tuple to get corresponding income prompts
            idx_income = income_tuple.index(income)
            income_indexed = income_list[idx_income]
        ###############################################################################################################################

        def show_messages(text):
            messages_str = [
                f"{_['content']}" for _ in st.session_state["messages"][1:] # Removed {_['role']}:
            ]
            text.text_area("Messages", value=str("\n".join(messages_str)), height=300)


        BASE_PROMPT = [{"role": "system", "content": ""}] # Fetches system

        if "messages" not in st.session_state:
            st.session_state["messages"] = BASE_PROMPT

        # Text input
        if not free_style_mode:
            value = "Enter your message here..."
        else:
            value = '''Generate a campaign email in 100 words for a bank with the following details:

Company: ICICI Bank
Product - Auto Loan
Role: Campaigns Writer
offer includes - no fees, 7% interest rate
tone - Encouraging
Link to apply - <URL>
offer valid for - 7 days'''
        txt = txt = st.text_area("Original Message", value=value, height = 200) #st.text_input("Prompt", value="Enter your message here...")
        # Output box
        text = st.empty()
        show_messages(text)

        # Final prompt
        if not free_style_mode:
            prompt = 'rephrase the following ' + campaign_type + 'concisely in less than ' + num_word + ' words '+ seg + age_indexed + income_indexed +' : "' + txt +'"' +"\n"
            # Create ICICI bank chatbot
            icici_chat_bot = ChatBot(num_char = num_char)
        else:
            prompt = txt
            icici_chat_bot = FreeStyle()

        if st.button("Send"):
            with st.spinner("Generating response..."):
                #st.session_state["messages"] += [{"role": "user", "content": prompt}] # Remove user's input message
                if not free_style_mode:
                    message_response = icici_chat_bot.get_response(prompt, cam_tone, role, product_indexed).strip() +"\n"
                else:
                    message_response = icici_chat_bot.get_response(prompt)
                st.session_state["messages"] += [
                    {"role": "system", "content": message_response}
                ]
                show_messages(text)

        if st.button("Clear"):
            st.session_state["messages"] = BASE_PROMPT
            show_messages(text)

    with tabs[1]:
        #st.subheader('Power Mode: Use ChatGPT to batch process campaign messages')
        # Dropdowns and sliders
        campaign_type = st.selectbox('Campaign Type ', campaign_list)
        product = st.selectbox('Product ', product_tuple)
        role = st.selectbox('Role ', role_tuple)
        age = st.selectbox('Age Band ', age_tuple)
        income = st.selectbox('Income Band ', income_tuple)
        cam_tone = st.selectbox('Campaign Tone ', tone_tuple)
        num_char = st.slider('Maximum Response Characters ', min_value=30, max_value=10000, step=1)
        num_word = str(int(round(num_char/6, 0)))
        # Use index of role_tuple to get corresponding role_content
        idx2 = role_tuple.index(role)
        content = role_content[idx2]
        # Use index of product_tuple to get corresponding product statements
        idx_prod = product_tuple.index(product)
        product_indexed = products[idx_prod]
        # Use index of age_tuple to get corresponding age prompts
        idx_age = age_tuple.index(age)
        age_indexed = products[idx_age]
        # Use index of income_tuple to get corresponding income prompts
        idx_income = income_tuple.index(income)
        income_indexed = income_list[idx_income]
        # File upload
        uploaded_file1 = st.file_uploader("Choose an Excel file", type="xlsx")

        if uploaded_file1 is not None:
            # Read the Excel file into a DataFrame
            df = pd.read_excel(uploaded_file1)

            # Get the column name
            col_name = df.columns[0]
            # Create empty list to store the rephrased df
            new_df_list = []

            # Loading animation
            with st.spinner("Generating response..."):
                new_col_list = [col_name + '_' + x for x in seg_tuple]
                result_list = [[col, []] for col in new_col_list]

                for index, row in df.iterrows():
                    # For each pyschographic segment
                    for x in seg_tuple:
                        # Create a new DataFrame to store the modified data
                        new_col = col_name + '_' + x
                        # Get the segments index
                        idx1 = seg_tuple.index(x)
                        # Get the prompt for that segment
                        seg = segments[idx1]
                        # Perform content designi for each row
                        # Final prompt
                        txt = row[col_name]
                        prompt = 'rephrase the following ' + campaign_type + 'concisely in less than ' + num_word + ' words '+ seg + age_indexed + income_indexed + ' : "' + txt +'"' +"\n"
                        # Create ICICI bank chatbot
                        icici_chat_bot = ChatBot(num_char = num_char)
                        # Generate response
                        message_response = icici_chat_bot.get_response(prompt, cam_tone, role, product_indexed).strip()
                        result_list[idx1][1].append(message_response)

            # Concatenate all the new df
            result_dict = {item[0]: item[1] for item in result_list}
            df_final = pd.DataFrame(result_dict)
            df_xlsx = to_excel(df_final)

            file_name = 'Batch Results' + '.xlsx'
            st.download_button(label='📥  Download Batch Results',
                                            data=df_xlsx ,
                                            file_name = file_name)
    with tabs[2]:
        ################################################### BLOB STORAGE ##############################################################
        # Client id, tenant id and client secret comes from Principal service
        client_id = 'bf85c804-31fe-4d5e-b226-749ca3796e75'
        tenant_id = '850aa78d-94e1-4bc6-9cf3-8c11b530701c'
        client_secret = 'W8Q8Q~zCZ3xvmgcENYljqLoP.OJqbbOvntoBqa.j'
        # Get this from Storage account > JSON view > BLOB
        account_url = 'https://storageicici.blob.core.windows.net/'

        # create a credential
        credentials = ClientSecretCredential(
            client_id = client_id,
            client_secret= client_secret,
            tenant_id= tenant_id)

        filename = 'Model Guidelines 1.xlsx'

        # Load the model guidelines File from Azure ADLS account
        def get_model_guidelines(filename, container_name = 'storagecommoncontainer', blob_name = 'Model Guidelines 1.xlsx'):
            '''Returns the vector store if exists else returns None'''
            # Create file path to store pickle file
            filepath = os.path.join(os.getcwd(), filename)
            # set client to access azure storage container
            blob_service_client = BlobServiceClient(account_url= account_url, credential= credentials)
            # get the container client
            container_client = blob_service_client.get_container_client(container=container_name)
            # List all blobs in the container client
            blob_list = container_client.list_blobs()
            # Loop runs only if blob exists
            for blob in blob_list:
                # Create blob client
                blob_client = container_client.get_blob_client(blob= blob_name)
                # Delete pickle file if already exists
                #delete_file(filepath)
                # Download the vectore sstore pickle file from Azure to local file path
                with open(filepath, "wb") as download_file:
                    download_file.write(blob_client.download_blob().readall())
                    #data = blob_client.download_blob().readall()
            # Unpickle the local pickle file to get the vector store
            with open(filepath, 'rb') as file:
                df_model_guide = pd.read_excel(filepath, sheet_name = 2)
            return df_model_guide

        ################################################### DROP DOWNS ##############################################################
        # Select Model - Defaults to DALL-E-3
        #model_tuple = ('DALL-E-3', 'DALL-E-2')
        #model = st.selectbox('Select Model', model_tuple)
        #model = 'DALL-E-3'
        model_tuple =('DALL-E-3', 'Segmind-sdxl')
        model = st.selectbox('Select Model', model_tuple)

        df_model_guide = get_model_guidelines(filename)
        df_model_guide.dropna(inplace = True)

        #if model == 'DALL-E-2':
            #size_tuple = ('256x256', '512x512' , '1024x1024')
            #size = st.selectbox('Image Size', size_tuple, index = 2) # Index to select tghe default value for the select box
            #num_images = st.slider('Number of Images', min_value=1, max_value=3, step=1)

        #model Dall-E-3
        if model == 'DALL-E-3':
            # Image Resolution
            size_tuple = ('1024x1024', '1024x1792', '1792x1024')
            size = st.selectbox('Image Size', size_tuple, index = 0)

            # Image quality
            quality_tuple = ('Standard', 'High Definition')
            quality_list = ['standard', 'hd']
            quality = st.selectbox('Image Quality', quality_tuple, index = 0)
            idx_quality = quality_tuple.index(quality)
            quality_value = quality_list[idx_quality]

            # Enable Model Guidelines
            enable_model_guidelines = st.checkbox("Enable Model Guidelines", key=None)

            if enable_model_guidelines:
                # Image cuts
                Model_tuple = tuple(df_model_guide['Model Cuts'].unique())
                Model_Cuts = st.selectbox('Model Cuts', Model_tuple, index = 0)

                #category_tuple = tuple(df_model_guide['Category'].unique())
                #category = st.selectbox('Category', category_tuple, index = 0)

                #age_band_tuple = tuple(df_model_guide['Age Band'].unique())
                #age_band = st.selectbox('Age Band', age_band_tuple, index = 0)

                #wear_band_tuple = tuple(df_model_guide['Wear'].unique())
                #wear = st.selectbox('Wear', wear_band_tuple, index = 0)

                filtered_df = df_model_guide[(df_model_guide['Model Cuts'] == Model_Cuts) ] # & (df_model_guide['Category'] == category) & (df_model_guide['Age Band'] == age_band) & (df_model_guide['Wear'] == wear)]
                guidelines = filtered_df['Final Guidelines'].iloc[0]


            # Number of Images
            num_images = st.slider('Number of Images', min_value=1, max_value=3, step=1)
            
        
        if model == 'Segmind-sdxl':
            # Image resolution in segmind is fixed to 1024x1024
            size_tuple = ('1024x1024','not availiable')
            size = st.selectbox('Image Size',size_tuple,index = 0)


            # Image quality
            quality_tuple = ('Standard', 'High Definition')
            quality_list = ['standard', 'hd']
            quality = st.selectbox('Image Quality', quality_tuple, index = 0)
            idx_quality = quality_tuple.index(quality)
            quality_value = quality_list[idx_quality]

            # Enable Model Guidelines
            enable_model_guidelines = st.checkbox("Enable Model Guidelines", key=None)

            if enable_model_guidelines:
                # Image cuts
                Model_tuple = tuple(df_model_guide['Model Cuts'].unique())
                Model_Cuts = st.selectbox('Model Cuts', Model_tuple, index = 0)

                #category_tuple = tuple(df_model_guide['Category'].unique())
                #category = st.selectbox('Category', category_tuple, index = 0)

                #age_band_tuple = tuple(df_model_guide['Age Band'].unique())
                #age_band = st.selectbox('Age Band', age_band_tuple, index = 0)

                #wear_band_tuple = tuple(df_model_guide['Wear'].unique())
                #wear = st.selectbox('Wear', wear_band_tuple, index = 0)

                filtered_df = df_model_guide[(df_model_guide['Model Cuts'] == Model_Cuts) ] # & (df_model_guide['Category'] == category) & (df_model_guide['Age Band'] == age_band) & (df_model_guide['Wear'] == wear)]
                guidelines = filtered_df['Final Guidelines'].iloc[0]


            # Number of Images
            num_images = st.slider('Number of Images', min_value=1, max_value=3, step=1)


        ################################################### IMAGE GENERATION CLASSES ##############################################################
        # Class for generating the images using DALL-E-2
        class ImageGeneratorDalle2:
            def __init__(self, prompt, num_images, size):
                self.prompt = prompt
                self.num_images = num_images
                self.size = size
                self.image_urls = []

            def generate_images(self):
                response = openai.Image.create(
                    prompt=self.prompt,
                    n=self.num_images,
                    size=self.size
                )
                self.image_urls = [response['data'][i]['url'] for i in range(self.num_images)]

            def display_images(self):
                for url in self.image_urls:
                    st.image(url)

        # Class for generating the images using Segmind-sdxl
        class ImageGeneratorSegmind:
         
            def __init__(self, prompt, num_images, size):
                self.prompt = prompt
                self.num_images = num_images
                self.size = size
                self.image_urls = []
                self.quality = quality_value 

            def generate_request_payload(self,prompt,seed, num_samples):

                return {

                  "prompt": prompt,

                  #"negative_prompt": "ugly, tiling, poorly drawn hands, poorly drawn feet, poorly drawn face, out of frame, , disfigured, deformed, body out of frame, blurry, bad anatomy, blurred, watermark, grainy, signature, cut off, draft",

                  "negative_prompt": "poorly drawn eyes,extra limbs, poorly drawn hands finger,extra hand,multiple fingers,bad eye,multiple hands,poorly drawn hands,bad men hair,missing hand,poorly drawn feet, poorly drawn face, out of frame, extra limbs, disfigured, deformed eyes, deformed body out of frame,bad anatomy,bad fingers, bad hands, missing fingers, blurry, bad anatomy, blurred, watermark",

                  "style": "base",

                  "samples": num_samples,

                  "scheduler": "UniPC",

                  "num_inference_steps": 25,

                  "guidance_scale": 8,

                  "strength": 0.2,

                  "seed": seed,

                  "img_width": 1024,

                  "img_height": 1024,

                  "refiner": True,

                  "high_noise_fraction": 0.8,

                  "base64": False

                }

            def generate_images(self):
                api_keys = ["SG_b6e45027b1ca6162", "SG_e3a7c58902a8f62f","SG_a7acf7af520ad30a","SG_4fbd24b92b397239"]
                #url = "https://api.segmind.com/v1/sdxl1.0-txt2img"
                url = "https://api.segmind.com/v1/sdxl1.0-realvis"
                # Randomly select an API key

                api_key = random.choice(api_keys)

                # Number of samples

                num_samples = self.num_images

                # Generate seeds based on the number of samples

                seeds = [random.randint(0, 999999) for _ in range(num_samples)]

                self.images=[]
                for seed in seeds:

                    response = requests.post(url, json=self.generate_request_payload(self.prompt,seed, num_samples), headers={'x-api-key': api_key})

                    if response.status_code == 200:

                        # Check if the response contains content

                        if response.content:
                        
                            image_data = io.BytesIO(response.content)

                            #display(image)
                            image = Image.open(image_data)
                            self.images.append(image)

                          
                              

                        else:

                            print("Error: Response does not contain any content")

                    else:

                        print(f"Error with API key {api_key}:", response.text)

            
            def display_images(self):
                for image in self.images:
                    st.image(image, caption='Generated Image', use_column_width=True)





        # Class for generating the images using DALL-E-3
        class ImageGeneratorDalle3:
            def __init__(self, prompt, num_images, size):
                self.prompt = prompt
                self.num_images = num_images
                self.size = size
                self.image_urls = []
                self.quality = quality_value

            def generate_images(self):
                client = OpenAI(api_key = openai.api_key)
                for i in range(self.num_images):
                    response = client.images.generate(
                    model="dall-e-3",
                    prompt=self.prompt,
                    size=self.size,
                    quality=self.quality,
                    n=1
                    )
                    image_url = response.data[0].url
                    self.image_urls.append(image_url)

            def display_images(self):
                for url in self.image_urls:
                    st.image(url)

        ################################################### PROMPT BOX + ENGINEERING ##############################################################

        prompt = st.text_area("Enter a detailed description about the image 🙋",height=100)

        if enable_model_guidelines:
            final_prompt = '''Use the following model/person guidelines followed by final image prompt for generating images.
            **MODEL GUIDELINES**
            Nationality - Indian Origin. Preferably should represent a pan Indian look.
            Model Cuts - {0}
            {1}

            **FINAL PROMPT FOR GENERATING IMAGE WHILE CONSIDERING THE ABOVE MODEL GUIDELINES**
            {2}. There should be no text inside the image.

            '''.format(Model_Cuts, guidelines, prompt)
        else:
            final_prompt = prompt


        image_button = st.button("Generate Image 🚀")
        clear_button = st.button("Clear Image")

        if image_button and prompt.strip() != "":
            with st.spinner("Loading...💫"):
                # Genrate and display images using Image Generator Class
                if model == 'DALL-E-2':
                    generator = ImageGeneratorDalle2(final_prompt, num_images, size)
                    generator.generate_images()
                    generator.display_images()

                if model =='Segmind-sdxl':
                    generator = ImageGeneratorSegmind(final_prompt, num_images, size)
                    generator.generate_images()
                    generator.display_images()
                    
                if model == 'DALL-E-3':
                    generator = ImageGeneratorDalle3(final_prompt, num_images, size)
                    generator.generate_images()
                    generator.display_images()

        elif clear_button:
            pass
        else:
            st.warning("Please enter something! ⚠")
