"""
Created on Mon Jul 11 15:16:27 2022

@author: piotr.janczewski
@author: justyna.zbiegien
"""

import os
import openai
import pandas as pd
import datetime as dt
import streamlit as st
from pptx import Presentation

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

CV_path = acn_path + "/Accenture/Warsaw Analytics - Documents/01_CVs/"
CV_flnm = "Warsaw_Analytics_FY23_template.pptx"
CV_file = CV_path + CV_flnm

AV_path = acn_path + "/Accenture/Staffing & productivity AI group - General/Dashboard/"
AV_flnm = "myScheduling_People_Extract.xlsx"
AV_file = AV_path + AV_flnm

LCR_flnm = "Staffing_data - LCR calc.xlsx"
LCR_file = AV_path + LCR_flnm

Promo_path = acn_path + "/Accenture/AI Executives Warsaw - General/"
Promo_flnm = "AI Ind Hub - promo slides.pptx"
Promo_file = Promo_path + Promo_flnm

App_path = acn_path + "/Desktop/genAI/CVapp"
Sel_txt = CV_path + '/CVapp/sel_list.txt'

dest_path = acn_path + "/Desktop/"

##############
# Dictionaries & initial parameters
##############

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

filtered_df = pd.DataFrame()

##############
# Functions
##############

# Scraping the pptx to produce a table with slidenums, names and positions
def scrap_CVs(CVprs):
    CVprs = Presentation(open(CV_file, "rb"))
    shape_list = []

    for slide in CVprs.slides:

        sld_nm = CVprs.slides.index(slide) + 1
        sldnt = slide.notes_slide
        if sldnt.notes_text_frame:
            sldnt_text = sldnt.notes_text_frame.text

            # Check if the shape is a text shape and if it has text
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

        # Map shape names to sections using shape_name_dict
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
    slides_to_remove = [slide for slide in presentation.slides if str(slide.slide_id) not in keep_slides_ids]

    # Remove slides
    for slide in slides_to_remove:
        presentation.slides.remove(slide)

    return presentation

def create_presentation(filtered_df, presentation, output_path):
    print('sld_nm list')
    print(filtered_df['sld_nm'].to_list())
    keep_slides_ids = filtered_df['sld_nm'].astype(str)
    print("keep_slides_ids w create_presentation")
    print(keep_slides_ids)
    presentation = remove_unwanted_slides(presentation, keep_slides_ids)
    presentation.save(output_path)

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
        st.session_state.filtered_df = All_df[(All_df['Select'] == True) & (All_df['AV'] ==1)]
        filtered_df = st.session_state.filtered_df
       
        print("======================================")
        print(filtered_df['sld_nm'].to_list())

        with st.expander("Filtered people list"):
            for index, row in filtered_df.iterrows():
                st.text(f"{row['Worker']} - {row['Dept']} - Level {row['Level']} - AV {row['AVweeks']}")

        filtered_df.to_csv("filtered_df.csv")

def filter_people(seniority_checks, person_checks, kwd_inp, dpt_DS, dpt_DE, dpt_Oth, av_sl):
    print('Success!')
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
    All_df['Availability week num'] = All_df['First Availability Date'].dt.isocalendar().week
    All_df['AVweeks'] = All_df['Availability week num'] - currweek
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

def final_export(filtered_df, CVprs):
    # Final export 

    dest = st.text_input("Enter the directory path to save the file:", acn_path)
    out_fn = st.text_input("Output file name", "CVs_free.pptx")
    export_button = st.button("Export all slides for the filtered people list")
     
    if export_button:
        filtered_df = pd.read_csv("filtered_df.csv")
        print("======================================")
        print("Tuż po przycisku eksport")
        print(filtered_df['sld_nm'].to_list())
        if not out_fn:
            out_fn = "AI Ind Hub CVs for .pptx"  
        output_path = f"{dest}/{out_fn if out_fn.endswith('.pptx') else out_fn + '.pptx'}"
        create_presentation(filtered_df, CVprs, output_path)
        st.success(f"Exported successfully to {output_path}")

    st.markdown("For help, visit [YouTube](https://www.youtube.com/watch?v=WNnzw90vxrE)")


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

initial_selection(All_df, shapes_df)

final_export(filtered_df, CVprs)