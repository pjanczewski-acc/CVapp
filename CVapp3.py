"""
Created on Mon Jul 11 15:16:27 2022

@author: piotr.janczewski
@author: justyna.zbiegien
"""

import os
import uuid
import json
#import openai
import pandas as pd
import datetime as dt
import streamlit as st
import unidecode as ud
import copy
from pptx.oxml.xmlchemy import OxmlElement
from datetime import datetime
from os.path import exists
from pptx import Presentation
from unidecode import unidecode

st.set_page_config(page_title="AI Industry Hub CV select",
                   # page_icon=folder + "/CVapp/images/favicon_accenture.png",
                   layout="wide",
                   initial_sidebar_state="expanded"
                   )


##############
# Input file locations and names
##############

acn_login = os.getlogin()
acn_path = "C:/Users/" + acn_login

CV_path = acn_path + "/OneDrive - Accenture/docs/01_CVs/"
CV_flnm = "Warsaw_Analytics_FY23_template.pptx"
CV_file = CV_path + CV_flnm

AV_path = acn_path + "/OneDrive - Accenture/Documents/moje dokumenty/CVapp/source samples/"
AV_flnm = "myScheduling_People_Extract.xlsx"
AV_file = AV_path + AV_flnm

LCR_flnm = "Staffing_data - LCR calc.xlsx"
LCR_file = AV_path + LCR_flnm

# Promo_path = acn_path + "/OneDrive - Accenture/Documents/moje dokumenty/CVapp/source samples/"
# Promo_flnm = "Promo_template.pptx"
# Promo_file = Promo_path + Promo_flnm

# App_path = acn_path + "/OneDrive - Accenture/Documents/moje dokumenty/CVapp/"
# Sel_txt = CV_path + '/CVapp/sel_list.txt'

dest_path = acn_path + "/Desktop/"

##############
# Dictionaries & initial parameters
##############

if 'filtered_df' not in st.session_state:
    st.session_state['filtered_df'] = pd.DataFrame()

if 'first_filtered' not in st.session_state:
    st.session_state['first_filtered'] = pd.DataFrame()


# Dictionary of shape names
shape_name_dict = {'Text Placeholder 1': 'About me',
                   'Text Placeholder 2': 'Industry experience',
                   'Text Placeholder 3': 'Name',
                   'Text Placeholder 4': 'Level & role',
                   'Text Placeholder 6': 'Education',
                   'Text Placeholder 7': 'Key skills',
                   'Text Placeholder 8': 'Languages',
                   'Text Placeholder 9': 'Relevant project experience'}

# Dictionary of titles per level
lvl_dict = {
    "13-New Associate": 13,
    "12-Associate": 12,
    "11-Analyst": 11,
    "10-Senior Analyst": 10,
    "9-Team Lead/Consultant": 9,
    "8-Associate Manager": 8,
    "7-Manager": 7,
    "7-Digital Data Innovation Senior Principal": 7,
    "6-Senior Manager": 6,
    "5-Associate Director": 5,
    "5-Analytics Principal Director": 5,
    "4-Accenture leadership": 4,
    "MD": 4
}
sen_vars = {}
for i in range(4, 13):
    sen_vars[f'sen_{i}'] = False

dpt_DS = False
dpt_DE = False
dpt_Oth = False

av_sl = 0

##############
# Functions
##############

# Scraping the pptx to produce a table with slidenums, names and positions
def scrap_CVs(CVprs):
   # CVprs.SectionList.RemoveAt(1)
    CVprs = Presentation(open(CV_file, "rb"))

    
    shape_list = []
    skip_slides = [1, 2]  # Slajdy do pominięcia

    for slide in CVprs.slides:
        sld_nm = CVprs.slides.index(slide) + 1

        if sld_nm in skip_slides:
            continue  # Pomijamy przetwarzanie tego slajdu

        sldnt = slide.notes_slide
        if sldnt.notes_text_frame:
            sldnt_text = sldnt.notes_text_frame.text

        for shape in slide.shapes:
            if shape.has_text_frame:
                text = "\n".join([paragraph.text for paragraph in shape.text_frame.paragraphs])

            if "4" in shape.name:
                if "data scien" in text.lower():
                    dept = "Data Science"
                elif "data engineer" in text.lower():
                    dept = "Data Engineering"
                else:
                    dept = "Other"

                shape_properties = {
                    'sld_nm': sld_nm,
                    'sld_nt': sldnt_text,
                    'name': shape.name,
                    'Dept': dept,
                    'text': text.strip()
                }

                shape_list.append(shape_properties)

        shapes_df = pd.DataFrame(shape_list)
        shapes_df['section'] = 'Other'
        shapes_df.loc[shapes_df['name'].str.startswith('Picture'), 'section'] = 'Picture'
        shapes_df.loc[shapes_df['name'].str.startswith('Text Pl'), 'section'] = shapes_df['name'].apply(
            lambda x: shape_name_dict.get(x, 'else'))

        names_df = shapes_df[['sld_nm', 'sld_nt', 'Dept']].drop_duplicates()
        names_df = names_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)

    return shapes_df, names_df, CVprs


# Function to fill missing Worker and Resource Name
def fill_missing_values(row):
    if pd.isna(row['EID']) and pd.notna(row['Enterprise ID']):
        row['EID'] = row['Enterprise ID']
    if pd.isna(row['Worker']) and pd.notna(row['EID']):
        row['Worker'] = row['EID'].title().replace('.', ' ')
    if pd.isna(row['Resource Name']):
        if pd.notna(row['Worker']):
            row['Resource Name'] = ', '.join(row['Worker'].split()[::-1])
    return row


# Loading input files
def load_inputs(AV_file, LCR_file, names_df):
    AV_df = pd.read_excel(AV_file, 'People_Extract', header=0)
    AV_df = AV_df[['Resource Name', 'Resource Email', 'First Availability Date']].dropna()
    AV_df['EID'] = AV_df['Resource Email'].str.split('@').str[0]
    AV_df['Resource Name'] = AV_df['Resource Name'].str.title()

    All_df = pd.read_excel(LCR_file, 'LCR', header=0)
    All_df = pd.merge(All_df, AV_df, left_on='Enterprise ID', right_on='EID', how='outer')
    All_df['EID'] = All_df['EID'].fillna(All_df['Enterprise ID'])
    All_df = pd.merge(All_df, names_df, left_on='EID', right_on='sld_nt', how='outer')

    # Apply the function to fill missing values
    All_df = All_df.apply(fill_missing_values, axis=1).dropna(subset=['EID', 'First Availability Date'], how='all')
    All_df = All_df.sort_values('Resource Name')
    All_df = All_df[['Worker', 'Resource Name', 'EID', 'sld_nm', 'Management Level', 'People Lead', 'LCR in $',
                     'First Availability Date', 'Dept']].sort_values('Resource Name')

    All_df = All_df.dropna(subset=['sld_nm'])
    return All_df

def remove_unwanted_slides(presentation, keep_slides_ids):
    """
    Usuwa slajdy, które nie znajdują się w podanym zbiorze identyfikatorów.
    Args:
    presentation (Presentation): Obiekt prezentacji.
    keep_slides_ids (set): Zbiór identyfikatorów slajdów do zachowania.
    """
    keep_slides_ids = set(map(lambda x: int(float(x)), keep_slides_ids)) # Konwersja na zestaw liczb całkowitych

    # Uzyskujemy dostęp do listy identyfikatorów slajdów
    slide_ids = presentation.slides._sldIdLst
    total_slides = len(slide_ids)
# Definiowanie zakresów slajdów, które mają być zawsze zachowane
    always_keep_first = 2
    always_keep_last = 3

    # Iterujemy w odwrotnej kolejności, aby indeksy nie były zakłócane po usunięciu
    for i in reversed(range(total_slides)):
        # Sprawdzamy, czy slajd jest poza zakresem pierwszych 'always_keep_first' i ostatnich 'always_keep_last' slajdów
        if (i >= always_keep_first and i < total_slides - (3 + always_keep_last)):
            # Usuwamy slajd tylko jeśli nie jest w zbiorze do zachowania
            if (i+1) not in keep_slides_ids:
                del slide_ids[i]

    return presentation


def create_presentation(filtered_df, presentation, output_path):
    keep_slides_ids = set(filtered_df['sld_nm'].astype(str))
    remove_unwanted_slides(presentation, keep_slides_ids)
    presentation.save(output_path)


def export_to_excel(df, filepath):
    df = df.rename(columns={
    'Resource Name': 'Name',
    'Management Level': 'Position Level'
    })

    # Konwersja i formatowanie daty
    if 'First Availability Date' in df.columns:
        df['First Availability Date'] = pd.to_datetime(df['First Availability Date']).dt.strftime('%d.%m.%Y')

    # Filtruj DataFrame, aby zawierał tylko potrzebne kolumny
    columns_to_export = ['Name', 'EID', 'People Lead', 'Position Level', 'First Availability Date', 'LCR in $']
    
    # Sprawdzanie, czy wszystkie wymagane kolumny są w DataFrame
    if all(column in df.columns for column in columns_to_export):
        fil_df = df[columns_to_export]
        print(fil_df)
        
        # Eksportowanie do pliku Excel
        with pd.ExcelWriter(filepath, engine='openpyxl', mode='w') as writer:
            fil_df.to_excel(writer, index=False)
        print(f"Data exported successfully to {filepath}")
    else:
        missing_columns = [column for column in columns_to_export if column not in df.columns]
        print(f"Missing columns in DataFrame: {missing_columns}")

# def keepSlides(keepID, prs):
#     # get slides to delete
#     ids = [x for x in range(1, len(prs.slides._sldIdLst) + 1) if x not in keepID]

#     # subset report
#     for i, slide in enumerate(prs.slides):
#         # create slide dict
#         id_dict = {slide.id: [i, slide.rId] for i, slide in enumerate(prs.slides._sldIdLst)}

#         # iterate thorugh indexes
#         if i + 1 in ids:
#             # get slide id
#             slide_id = slide.slide_id

#             # remove slide
#             prs.part.drop_rel(id_dict[slide_id][1])
#             del prs.slides._sldIdLst[id_dict[slide_id][0]]

#     return prs


# def kwdlookup(kwd_input, sld):
#     kwds = kwd_input.lower()
#     if 'and' in kwds:
#         kwd_lgc = 'and'
#         kwd_lst = kwds.split(" and")
#     elif 'or' in kwd_input:
#         kwd_lgc = 'or'
#         kwd_lst = kwds.split(" or")
#     else:
#         kwd_lgc = 'na'
#         kwd_lst = []
#         kwd_lst.append(kwds)
#     kwd_fnd = 0
#     for i, kwd in enumerate(kwd_lst):
#         for shape in CVprs.slides[sld].shapes:
#             if hasattr(shape, "text") and (kwd in shape.text.lower()):
#                 kwd_fnd = kwd_fnd + 1
#                 break
#     if kwd_lgc == 'and' and kwd_fnd == len(kwd_lst): kwd_fnd = 1
#     elif kwd_lgc == 'or' and kwd_fnd > 0: kwd_fnd = 1
#     elif kwd_lgc == 'na' and kwd_fnd == 1: kwd_fnd = 1
#     elif kwds == '': kwd_fnd = 1

#     return kwd_fnd


def initial_selection(All_df, shapes_df):
    st.markdown("Specify criteria to export one-slider CVs")
    st.write("")

    # with st.form("CV selection", clear_on_submit=False):

    av_sl = st.slider("Availability time (in weeks, counted" +
                        " since last Monday): 0=bench, 14=Over 3M",
                        min_value=0, max_value=14)
    st.write("")
    seniority_checks = {}
    with st.expander("Seniority level"):
        seniority_levels = [
            "4 - Accenture Leadership", 
            "5 - Associate Director", 
            "6 - Senior Manager",
            "7 - Manager", 
            "8 - Associate Manager", 
            "9 - Team Lead/Consultant",
            "10 - Senior Analyst", 
            "11 - Analyst", 
            "12 - Below Analyst"
        ]

        for level in seniority_levels:
            level_id = int(level.split(' - ')[0])
            seniority_checks[level_id] = st.checkbox(level)

    with st.expander("Department/profile"):
        dpt_DS = st.checkbox("Data Science")
        dpt_DE = st.checkbox("Data Engineering")
        dpt_Oth = st.checkbox("Other")

    with st.expander("Keywords"):
        st.markdown("<p style='font-size:12px;'>" +
                    "Use AND/OR (never both), eg.:<i>" +
                    " 'machine learning AND Azure AND risk advisory' or " +
                    " 'gcp OR Google Cloud Platform'</i></p>",
                    unsafe_allow_html=True)
        kwd_inp = st.text_input('Text to look up:')

    All_df.reset_index(drop=True, inplace=True)

    person_checks = {}
    with st.expander("Preliminary person list"):
        for index, Person in enumerate(All_df['Worker'], start=0):
            person_checks[index] = st.checkbox(f"{index} {Person}")

    # Premilinary listing before final export

    All_df['Select'] = False
    All_df['AV'] = 0

    listed = st.button("Filter people for final selection")
    if listed:
        filter_people(seniority_checks, person_checks, kwd_inp, dpt_DS, dpt_DE, dpt_Oth, av_sl)
        

    # Displaying people for final approval
    filtered_df = All_df[(All_df['Select'] == True) & (All_df['AV'] ==1)]

    if not filtered_df.empty:
        st.session_state['filtered_df'] = filtered_df
        st.session_state['first_filtered'] = filtered_df

    
    
    with st.expander("Filtered people list"):
        if not st.session_state['first_filtered'].empty:
            for index, row in st.session_state['first_filtered'].iterrows():
                checked=st.checkbox(f"{row['Worker']} - {row['Dept']} - Level {row['Level']} - AV {row['AVweeks']}", value=True)
                st.session_state['filtered_df'].loc[index, 'Select'] = checked

    if not st.session_state['filtered_df'].empty:
        st.session_state['filtered_df'] = st.session_state['filtered_df'][st.session_state['filtered_df']['Select'] == True]

    return st.session_state.get('filtered_df', pd.DataFrame()), seniority_checks, person_checks, kwd_inp, dpt_DS, dpt_DE, dpt_Oth, av_sl

def filter_people(seniority_checks, person_checks, kwd_inp, dpt_DS, dpt_DE, dpt_Oth, av_sl):
    active_filters_count =0
    # Counting active filters
    if any(seniority_checks.values()):
        active_filters_count += 1
    active_filters_count += (dpt_DS or dpt_DE or dpt_Oth) + bool(kwd_inp)

    # Adding Level
    All_df["Level"] = All_df['Management Level'].apply(lambda x: lvl_dict.get(x))

    # Adding availability
    All_df["Avlbl"] = 14
    currweek = dt.date.today().isocalendar().week
    
    # if pd.to_datetime(All_df['First Availability Date']) < datetime.now():
    #     All_df['Availability week num'] = 0
    #     All_df['AVweeks'] = 0
    # else:
    #     All_df['Availability week num'] = All_df['First Availability Date'].dt.isocalendar().week
    #     All_df['AVweeks'] = All_df['Availability week num'] - currweek

    # All_df['First Availability Date'] = pd.to_datetime(All_df['First Availability Date'], errors='coerce')

    # Sprawdzenie, czy każda data w kolumnie 'First Availability Date' jest mniejsza niż dzisiejsza data
    curr_date = pd.to_datetime(dt.datetime.now())
    curr_year = curr_date.year
  
    All_df['First Availability Date'] = pd.to_datetime(All_df['First Availability Date'], errors='coerce')

    # Sprawdzenie, czy rok daty jest wyższy niż obecny rok
    mask_future_year = All_df['First Availability Date'].dt.year > curr_year

    # Obliczenie tygodni do końca roku dla daty w przyszłym roku
    weeks_to_end_of_year = 52 - dt.date.today().isocalendar().week

    # Ustawienie wartości dla wierszy, które spełniają warunek przyszłego roku
    All_df.loc[mask_future_year, 'Availability week num'] = weeks_to_end_of_year + All_df.loc[mask_future_year, 'First Availability Date'].dt.isocalendar().week
    All_df.loc[mask_future_year, 'AVweeks'] = All_df.loc[mask_future_year, 'Availability week num'] - currweek

    # Obliczenie numeru tygodnia dla dat w bieżącym roku
    mask_current_year = All_df['First Availability Date'].dt.year == curr_year
    current_year_dates = All_df.loc[mask_current_year, 'First Availability Date']

    # Sprawdzenie, czy każda data w kolumnie 'First Availability Date' jest mniejsza niż dzisiejsza data
    mask = All_df['First Availability Date'] < pd.to_datetime(dt.datetime.now())

    # Ustawienie wartości dla wierszy, które spełniają warunek
    All_df.loc[mask, 'Availability week num'] = 0
    All_df.loc[mask, 'AVweeks'] = 0

    # Obliczenie numeru tygodnia dla dat, które nie spełniają warunku
    mask_current_year_and_future = mask_current_year & (current_year_dates >= pd.to_datetime(dt.datetime.now())) & ~mask_future_year
    All_df.loc[mask_current_year_and_future, 'Availability week num'] = current_year_dates.dt.isocalendar().week
    All_df.loc[mask_current_year_and_future, 'AVweeks'] = All_df.loc[mask_current_year_and_future, 'Availability week num'] - currweek

    
    All_df['AV'] = 0
    All_df.loc[All_df['AVweeks'] <= av_sl, 'AV'] = 1


    # Flitering criteria
    for index, row in All_df.iterrows():
        matched_criteria_count = 0

        # Counting criteria selected
        if row['Level'] in [lvl for lvl, checked in seniority_checks.items() if checked]:
                matched_criteria_count += 1

        if ((dpt_DS and 'sci' in row['Dept'].lower()) or
                (dpt_DE and 'eng' in row['Dept'].lower()) or
                (dpt_Oth and 'sci' not in row['Dept'].lower() and 'eng' not in row['Dept'].lower())):
            matched_criteria_count += 1

        # kwd_imp should be here

        All_df.loc[index, 'Select'] = matched_criteria_count == active_filters_count

        # Checking initial people list
        if person_checks.get(index):
            All_df.loc[index, 'Select'] = True
            All_df.loc[index, 'AV'] = 1

    
    

def final_export(filtered_df, CVprs, seniority_checks, person_checks, kwd_inp, dpt_DS, dpt_DE, dpt_Oth, av_sl):
    # Final export 

    dest = st.text_input("Enter the directory path to save the file:", "path/to/directory")
    out_fn = st.text_input("Output file name", "AI Ind Hub CVs for .pptx")
    export_button = st.button("Export all slides for the filtered people list")

    out_fn_excel = st.text_input("Output file name for Excel", "AI Ind Hub CVs table.xlsx")
    export_excel_button = st.button("Export filtered table data to Excel")
    
     
    if export_button:
        filtered_df = st.session_state.get('filtered_df', pd.DataFrame())

        if not filtered_df.empty:
            filtered_df = st.session_state['filtered_df']
            
            if not out_fn:
                out_fn = "AI Ind Hub CVs for .pptx"  
            output_path = f"{dest}/{out_fn if out_fn.endswith('.pptx') else out_fn + '.pptx'}"

            create_presentation(filtered_df, CVprs, output_path)
            st.success(f"Exported successfully to {output_path}")
        else:
            st.error("Please filter the data before exporting.")

    #st.write(st.session_state['filtered_df'])
    if export_excel_button:
        # Logika eksportu do Excela
        if 'filtered_df' in st.session_state and not st.session_state['filtered_df'].empty:
            filtered_df = st.session_state['filtered_df']
            if not out_fn_excel:
                out_fn = "AI Ind Hub CVs table.xlsx"  
            output_path = f"{dest}/{out_fn_excel if out_fn_excel.endswith('.xlsx') else out_fn_excel + '.xlsx'}"
            export_to_excel(filtered_df, output_path)
            st.success(f"Data exported successfully to {output_path}")
        else:
            st.error("Please filter the data before exporting.")

    st.markdown("For help, visit [YouTube](https://www.youtube.com/watch?v=WNnzw90vxrE)")

    st.info("Before using the application, please ensure that all data is up-to-date.")
    st.info("The application reads the employee's level based on the export file from MyScheduling. Also, make sure you have the latest CV presentations for all team members.")
    st.info("Note! If there is a CV for a particular person in the presentation but it does not exist in the export file from MyScheduling, you will not see that person when filtering individuals. However, you can manually select them in the Preliminary Person list section. Additionally, if a person does not exist on the MyScheduling list, their Level will be set to 'NaN'.")
##############
# PAGE SET UP
##############

hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            </style>
            """

shapes_df, names_df, CVprs = scrap_CVs(CV_file)
All_df = load_inputs(AV_file, LCR_file, names_df)

filtered_df, seniority_checks, person_checks, kwd_inp, dpt_DS, dpt_DE, dpt_Oth, av_sl = initial_selection(All_df, shapes_df)

final_export(filtered_df, CVprs, seniority_checks, person_checks, kwd_inp, dpt_DS, dpt_DE, dpt_Oth, av_sl)