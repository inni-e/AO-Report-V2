# Module for Creating Report 

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns # Version 0.11.2 required
from fpdf import FPDF # Version 2.3.1 required
from collections import Counter
from PyPDF2 import PdfMerger, PdfReader  # Updated imports

# selective calculator module
from selec_calc import selective_calc
from oc_calc import oc_calc

# PDF module
from pdf_module import PDF

# Plot module
import plots

# Helper functions
def flexi_extract(dataset, test_type, is_eng = False, ans_sheet = None):
    if is_eng:
        # EXAM VARIABLES
        if test_type == 'sttc':
            cloze_start = 15
            num_cloze = 6
            num_extracts = 10
        elif test_type == 'octt':
            cloze_start = 12
            num_cloze = 6
            num_extracts = 8
        else:
            cloze_start = 6
            num_cloze = 5
            num_extracts = 5

        # first extract normal indices
        normal_indices = np.array([])
        for i in range(11, len(dataset.columns)): 
            if i != len(dataset.columns)-1:
                if all(dataset.iloc[:, i] == 1) and not all(dataset.iloc[:, i+1] == 1):
                    normal_indices = np.append(normal_indices, i-1)
            
        normal_student = dataset.to_numpy()[:, normal_indices.astype(int)]
        
        # function to clean answer sheet
        clean_array = np.vectorize(lambda x: x.strip().upper())

        # extract cloze passage
        cloze_indices = np.arange(normal_indices[-1]+2, normal_indices[-1] + 2 + num_cloze)
        cloze_students = dataset.to_numpy()[:, cloze_indices.astype(int)]
        cloze_students[pd.isna(cloze_students)] = 'Z' # filler answer
        
        # extract comparing excerpts 
        extract_indices = np.arange(cloze_indices[-1]+4, cloze_indices[-1] + 4 + num_extracts)
        extract_students = dataset.to_numpy()[:, extract_indices.astype(int)]
        extract_students[pd.isna(extract_students)] = 'Z'

        unnormal_students = np.concatenate((cloze_students, extract_students), axis = 1)
        unnormal_students[pd.isna(unnormal_students)] = 'Z'
        unnormal_students = clean_array(unnormal_students)
        
        
        # get answers
        unnormal_ans = ans_sheet.iloc[(cloze_start-1):, 1]
        unnormal_students = np.array([unnormal_students[i, ] == unnormal_ans for i in range(0, len(unnormal_students))])
        unnormal_students = unnormal_students.astype(int)
        
        return(np.concatenate((normal_student, unnormal_students), axis = 1))
        
    else:
        indices = np.array([])
        for i in range(11, len(dataset.columns)):
            if i != len(dataset.columns)-1:
                if all(dataset.iloc[:, i] == 1) and not all(dataset.iloc[:, i+1] == 1):
                    indices = np.append(indices, i-1)
            else:
                indices = np.append(indices, i-1)
        ans_sheet = dataset.to_numpy()[:, indices.astype(int)]
        
    return(ans_sheet)  

def phone_ans_sheet(dataset):
    ans_sheet = dataset.iloc[:, np.arange(13, dataset.shape[1]+1, 3)]
    return(ans_sheet)

# Retrieve email from rollid
# def get_email(rollid, mark_file):
#     emails = [rollid['EMAILID'][np.where(rollid.NAME == name)[0]].to_string(index = False) for name in mark_file['Name']]
#     return(emails)

def get_email(rollid, mark_file):
    emails = ['']*len(mark_file)
    for i in range(len(mark_file)):
        id = mark_file.SID[i]
        if id in list(rollid.iloc[:, 0]):
            emails[i] = rollid.iloc[np.where(rollid.iloc[:, 0] == id)[0], 3].values[0]
        else:
            emails[i] = 'Missing'
    return(emails)


# Report Creation Class
class ReportCreation():
    def __init__(self, test_type, test_no, rollid, reading_data, maths_data, thinking_data, writing_data = None, data_type = 'hybrid'):
        ##### Parameter Values #######
        # test_type: sttc or octt
        # data_type: hybrid (default), inperson or online
        self.test_type = test_type[0:4]
        self.test_type2 = test_type
        self.test_no = test_no
        self.rollid = rollid
        self.reading_data = reading_data
        self.maths_data = maths_data
        self.thinking_data = thinking_data
        self.data_type = data_type
        
        if test_type != 'octt':
            self.writing_data = writing_data
        
    
    ############### Helper Function ##################
    # Create incomplete student dataframe to append to aggregate dataframe
    def generate_incomplete_df(self, incomplete_students, has_reading, has_maths, has_thinking, has_writing = None):
        incomplete_df = np.full([len(incomplete_students), self.agg_data.shape[1]], np.nan)
        incomplete_df = pd.DataFrame(incomplete_df, columns = self.agg_data.columns)
        
        ########### Add in student data for reading ############
        reading_ind = [i for i in range(len(incomplete_students)) if incomplete_students[i] in list(has_reading.SID)]
        
        # Name, email, reading mark, and reading ans sheet
        incomplete_df.iloc[reading_ind, 0] = has_reading.Name
        incomplete_df.iloc[reading_ind, [1, 2]] = has_reading.iloc[:, [-1, 4]]
        incomplete_df.iloc[reading_ind, -self.thinking_total-self.maths_total-self.reading_total:-self.thinking_total-self.maths_total] = has_reading.iloc[:, -1-self.reading_total:-1]
        
        ########## Add in student data for maths #############
        maths_ind = [i for i in range(len(incomplete_students)) if incomplete_students[i] in list(has_maths.SID)]
        
        # Email, maths mark, and maths ans sheet
        incomplete_df.iloc[maths_ind, 0] = has_maths.Name
        incomplete_df.iloc[maths_ind, [1, 3]] = has_maths.iloc[:, [-1, 4]]
        incomplete_df.iloc[maths_ind, -self.thinking_total-self.maths_total:-self.thinking_total] = has_maths.iloc[:, -1-self.maths_total:-1]
        
        ######### Add in student data for thinking ###############
        thinking_ind = [i for i in range(len(incomplete_students)) if incomplete_students[i] in list(has_thinking.SID)]
        
        # Email, thinking mark, and thinking ans sheet
        incomplete_df.iloc[thinking_ind, 0] = has_thinking.Name
        incomplete_df.iloc[thinking_ind, [1, 4]] = has_thinking.iloc[:, [-1, 4]]
        incomplete_df.iloc[thinking_ind, -self.thinking_total:] = has_thinking.iloc[:, -1-self.thinking_total:-1].astype('float64')
        
        if has_writing is not None:
            writing_ind = [i for i in range(len(incomplete_students)) if incomplete_students[i] in list(has_writing.SID)]
            
            # Email and writing mark
            incomplete_df.iloc[writing_ind, 0] = has_writing.Name
            incomplete_df.iloc[writing_ind, [1, 5]] = has_writing.iloc[:, [-1, 2]]
            
            # Add ranks
            incomplete_df.iloc[reading_ind, 7] = has_reading.ranking
            incomplete_df.iloc[maths_ind, 8] = has_maths.ranking
            incomplete_df.iloc[thinking_ind, 9] = has_thinking.ranking
            incomplete_df.iloc[writing_ind, 10] = has_writing.ranking
        
        else:
            # Add ranks for tests without writing
            incomplete_df.iloc[reading_ind, 6] = has_reading.ranking
            incomplete_df.iloc[maths_ind, 7] = has_maths.ranking
            incomplete_df.iloc[thinking_ind, 8] = has_thinking.ranking
        
        return(incomplete_df)
    
    
    def prepare(self):    
        if self.data_type == 'hybrid' or self.data_type == 'phone':
            # Hybrid will already be in phone dataset style
            self.num_reading = len(self.reading_data)
            self.num_maths = len(self.maths_data)
            self.num_thinking = len(self.thinking_data)
                            
            # Record total marks
            self.reading_total = int(np.sum(self.reading_data.iloc[0, 7:10]))
            self.maths_total = int(np.sum(self.maths_data.iloc[0, 7:10]))
            self.thinking_total = int(np.sum(self.thinking_data.iloc[0, 7:10]))
            
            # Ranking Columns
            # Creating list of marks
            reading_marks = -np.sort(-self.reading_data['Total Marks'])
            maths_marks = -np.sort(-self.maths_data['Total Marks'])
            thinking_marks = -np.sort(-self.thinking_data['Total Marks'])
            
            # ranking column
            self.reading_data['ranking'] = [int(np.where(reading_marks == mark)[0][0])+1 for mark in self.reading_data['Total Marks']]
            self.maths_data['ranking'] = [int(np.where(maths_marks == mark)[0][0])+1 for mark in self.maths_data['Total Marks']]
            self.thinking_data['ranking'] = [int(np.where(thinking_marks == mark)[0][0])+1 for mark in self.thinking_data['Total Marks']]
            
            # Add answer sheets to data
            # Reading comprehension
            r_sheet = phone_ans_sheet(self.reading_data)
            r_sheet.columns = ['rQ{}'.format(i) for i in range(1, 1+self.reading_total)]
            self.reading_data = self.reading_data.join(r_sheet)
            
            # Mathematical Reasoning                 
            m_sheet = phone_ans_sheet(self.maths_data)
            m_sheet.columns = ['mQ{}'.format(i) for i in range(1, 1+self.maths_total)]
            self.maths_data = self.maths_data.join(m_sheet)
            
            # Thinking skills
            t_sheet = phone_ans_sheet(self.thinking_data)
            t_sheet.columns = ['tQ{}'.format(i) for i in range(1, 1+self.thinking_total)]
            self.thinking_data = self.thinking_data.join(t_sheet)
            
            
            # Grab Emails
            self.reading_data['Email'] = get_email(self.rollid, self.reading_data)
            self.maths_data['Email'] = get_email(self.rollid, self.maths_data)
            self.thinking_data['Email'] = get_email(self.rollid, self.thinking_data)
            
            # Rename columns to be consistent with other data types
            self.reading_data = self.reading_data.rename(columns = {'Total Marks' : 'Points'})      
            self.maths_data = self.maths_data.rename(columns = {'Total Marks' : 'Points'})
            self.thinking_data = self.thinking_data.rename(columns = {'Total Marks' : 'Points'})
            
            # Sort students by alphabetical order
            self.reading_data = self.reading_data.sort_values('SID')
            self.maths_data = self.maths_data.sort_values('SID')
            self.thinking_data = self.thinking_data.sort_values('SID')
            
            if self.test_type != 'octt':
                # Repeat above processes for writing
                self.num_writing = len(self.writing_data)
                self.writing_total = 25 if self.test_type2 == 'sttc' or self.test_type2 == 'wemtsel' else 20
                writing_marks = -np.sort(-self.writing_data.iloc[:, 2])
                self.writing_data['ranking'] = [int(np.where(writing_marks == mark)[0][0])+1 for mark in self.writing_data.iloc[:, 2]]
                self.writing_data['Email'] = get_email(self.rollid, self.writing_data) 
                self.writing_data = self.writing_data.sort_values('SID')
        
        elif self.data_type == 'online':
            # Number of Students
            self.num_reading = len(self.reading_data)
            self.num_maths = len(self.maths_data)
            self.num_thinking = len(self.thinking_data)
            
            # Record Total Marks
            self.reading_total = self.reading_data.iloc[0, 5]
            self.maths_total = self.maths_data.iloc[0, 5]
            self.thinking_total = self.thinking_data.iloc[0, 5]
            
            # Ranking Columns
            # Creating list of marks
            reading_marks = -np.sort(-self.reading_data['Points'])
            maths_marks = -np.sort(-self.maths_data['Points'])
            thinking_marks = -np.sort(-self.thinking_data['Points'])
            
            # ranking column
            self.reading_data['ranking'] = [int(np.where(reading_marks == mark)[0][0])+1 for mark in self.reading_data['Points']]
            self.maths_data['ranking'] = [int(np.where(maths_marks == mark)[0][0])+1 for mark in self.maths_data['Points']]
            self.thinking_data['ranking'] = [int(np.where(thinking_marks == mark)[0][0])+1 for mark in self.thinking_data['Points']]
            
            # Add answer sheets to data
            reading_ans_sheet = pd.read_excel('reading_q_types.xlsx')
            r_sheet = flexi_extract(self.reading_data, self.test_type, is_eng = True, ans_sheet = reading_ans_sheet)
            r_sheet.columns = ['rQ{}'.format(i) for i in range(1, 1+self.reading_total)]
            self.reading_data = self.reading_data.join(r_sheet)
            
            # Mathematical Reasoning                 
            m_sheet = flexi_extract(self.maths_data, self.test_type)
            m_sheet.columns = ['mQ{}'.format(i) for i in range(1, 1+self.maths_total)]
            self.maths_data = self.maths_data.join(m_sheet)
            
            # Thinking skills
            t_sheet = flexi_extract(self.thinking_data, self.test_type)
            t_sheet.columns = ['tQ{}'.format(i) for i in range(1, 1+self.thinking_total)]
            self.thinking_data = self.thinking_data.join(t_sheet)
            
            # Sort
            self.reading_data = self.reading_data.sort_values('Last name')
            self.maths_data = self.maths_data.sort_values('Last name')
            self.thinking_data = self.thinking_data.sort_values('Last name')
            
            # Rename
            self.reading_data = self.reading_data.rename(columns = {'First name':'Name', 'Last name':'SID'})
            self.maths_data = self.maths_data.rename(columns = {'First name':'Name', 'Last name':'SID'})
            self.thinking_data = self.thinking_data.rename(columns = {'First name':'Name', 'Last name':'SID'})
                        
            if self.test_type != 'octt':
                # Repeat above process but for writing
                self.num_writing = len(self.writing_data)
                self.writing_total = 25 if self.test_type2 == 'sttc' or self.test_type2 == 'wemtsel' else 20
                writing_marks = -np.sort(-self.writing_data.iloc[:, 4])
                self.writing_data['ranking'] = [int(np.where(writing_marks == mark)[0][0])+1 for mark in self.writing_data.iloc[:, 4]]
                self.writing_data = self.writing_data.sort_values('Last name')
                self.writing_data = self.writing_data.rename(columns = {'First name':'Name', 'Last name':'SID'})


    def aggregate_data(self):
        # Separate by complete and incomplete students
        if self.test_type == 'octt':
            complete_students = set(self.reading_data.SID).intersection(self.maths_data.SID).intersection(self.thinking_data.SID)
            self.num_complete = len(complete_students)
            
            # datasets for complete students
            complete_reading = self.reading_data.iloc[[id in complete_students for id in self.reading_data.SID],].reset_index(drop = True)
            complete_maths = self.maths_data.iloc[[id in complete_students for id in self.maths_data.SID], ].reset_index(drop = True)
            complete_thinking = self.thinking_data.iloc[[id in complete_students for id in self.thinking_data.SID], ].reset_index(drop = True)
            
            # Create aggregate data for complete students
            
            # mark info
            self.agg_data = {
                'Name' : complete_reading.Name,
                'email' : complete_reading.Email, 
                'reading_mark' : complete_reading.Points,
                'maths_mark' : complete_maths.Points,
                'thinking_mark' : complete_thinking.Points
            }
            self.agg_data = pd.DataFrame(self.agg_data)
            
            # compute OC mark
            self.agg_data['oc_mark'] = [oc_calc(self.agg_data.iloc[i, 2:]) for i in range(len(self.agg_data))]

            # rank info
            self.agg_data['reading_rank'] = complete_reading.ranking
            self.agg_data['maths_rank'] = complete_maths.ranking
            self.agg_data['thinking_rank'] = complete_thinking.ranking

            # overall rank
            oc_sorted = -np.sort(-self.agg_data.oc_mark)
            self.agg_data['overall_rank'] = [np.where(oc_sorted == mark)[0][0]+1 for mark in self.agg_data.oc_mark]

            # add on answer sheets
            self.agg_data = self.agg_data.join(complete_reading.iloc[:, -1-self.reading_total:-1]).join(complete_maths.iloc[:, -1-self.maths_total:-1]).join(complete_thinking.iloc[:, -1-self.thinking_total:-1])
            
            
            # Deal with incomplete students
            has_reading = self.reading_data.iloc[[not id in complete_students for id in self.reading_data.SID], ].reset_index(drop = True)
            has_maths = self.maths_data.iloc[[not id in complete_students for id in self.maths_data.SID], ].reset_index(drop = True)
            has_thinking = self.thinking_data.iloc[[not id in complete_students for id in self.thinking_data.SID], ].reset_index(drop = True)
            
            # list of incomplete students
            incomplete_students = np.concatenate([has_reading.SID, has_maths.SID, has_thinking.SID])
            incomplete_students = np.unique(incomplete_students)
            
            # incomplete student infos
            incomplete_df = self.generate_incomplete_df(incomplete_students, has_reading, has_maths, has_thinking)
            
            # combine complete with incomplete df
            self.agg_data = pd.concat([self.agg_data, incomplete_df], ignore_index = True)
            self.agg_data.iloc[:, 10:] = self.agg_data.iloc[:, 10:].astype('float64')
            return(self.agg_data)

            
        elif self.test_type == 'sttc':
            complete_students = set(self.reading_data.SID).intersection(self.maths_data.SID).intersection(self.thinking_data.SID).intersection(self.writing_data.SID)
            self.num_complete = len(complete_students)
            
            # datasets for complete students
            complete_reading = self.reading_data.iloc[[id in complete_students for id in self.reading_data.SID],].reset_index(drop = True)
            complete_maths = self.maths_data.iloc[[id in complete_students for id in self.maths_data.SID], ].reset_index(drop = True)
            complete_thinking = self.thinking_data.iloc[[id in complete_students for id in self.thinking_data.SID], ].reset_index(drop = True)
            complete_writing = self.writing_data.iloc[[id in complete_students for id in self.writing_data.SID], ].reset_index(drop = True)
            
            # Create aggregate data for complete students
            
            # Mark info
            self.agg_data = {
                'Name' : complete_reading.Name,
                'email' : complete_reading.Email, 
                'reading_mark' : complete_reading.Points,
                'maths_mark' : complete_maths.Points,
                'thinking_mark' : complete_thinking.Points,
                'writing_mark' : complete_writing.Points
            }
            self.agg_data = pd.DataFrame(self.agg_data)
            
            # compute selective mark
            self.agg_data['selec_mark'] = [selective_calc(self.agg_data.iloc[i, 2:]) for i in range(len(self.agg_data))]

            # Rank info
            self.agg_data['reading_rank'] = complete_reading.ranking
            self.agg_data['maths_rank'] = complete_maths.ranking
            self.agg_data['thinking_rank'] = complete_thinking.ranking
            self.agg_data['writing_rank'] = complete_writing.ranking
            
            # Overall ranking
            selective_sorted = -np.sort(-self.agg_data.selec_mark)
            self.agg_data['overall_rank'] = [np.where(selective_sorted == mark)[0][0]+1 for mark in self.agg_data.selec_mark]
            
            # add on answer sheets
            self.agg_data = self.agg_data.join(complete_reading.iloc[:, -1-self.reading_total:-1]).join(complete_maths.iloc[:, -1-self.maths_total:-1]).join(complete_thinking.iloc[:, -1-self.thinking_total:-1])
            
            
            # Deal with incomplete students
            has_reading = self.reading_data.iloc[[not student in complete_students for student in self.reading_data.SID], ].reset_index(drop = True)
            has_maths = self.maths_data.iloc[[not student in complete_students for student in self.maths_data.SID], ].reset_index(drop = True)
            has_thinking = self.thinking_data.iloc[[not student in complete_students for student in self.thinking_data.SID], ].reset_index(drop = True)
            has_writing = self.writing_data.iloc[[not student in complete_students for student in self.writing_data.SID], ].reset_index(drop = True)
            
            # list of incomplete students
            incomplete_students = np.concatenate([has_reading.SID, has_maths.SID, has_thinking.SID, has_writing.SID])
            incomplete_students = np.unique(incomplete_students)
                        
            # incomplete student infos
            incomplete_df = self.generate_incomplete_df(incomplete_students, has_reading, has_maths, has_thinking, has_writing)

            
            # Combine incomplete dataframe with complete
            self.agg_data = pd.concat([self.agg_data, incomplete_df], ignore_index = True)
            self.agg_data.iloc[:, 12:] = self.agg_data.iloc[:, 12:].astype('float64')
            return(self.agg_data)
        
        else: # wemt
            complete_students = set(self.reading_data.SID).intersection(self.maths_data.SID).intersection(self.thinking_data.SID).intersection(self.writing_data.SID)
            self.num_complete = len(complete_students)
            
            # datasets for complete students
            complete_reading = self.reading_data.iloc[[id in complete_students for id in self.reading_data.SID],].reset_index(drop = True)
            complete_maths = self.maths_data.iloc[[id in complete_students for id in self.maths_data.SID], ].reset_index(drop = True)
            complete_thinking = self.thinking_data.iloc[[id in complete_students for id in self.thinking_data.SID], ].reset_index(drop = True)
            complete_writing = self.writing_data.iloc[[id in complete_students for id in self.writing_data.SID], ].reset_index(drop = True)
            
            # Create aggregate data for complete students
            
            # Mark info
            self.agg_data = {
                'Name' : complete_reading.Name,
                'email' : complete_reading.Email, 
                'reading_mark' : complete_reading.Points,
                'maths_mark' : complete_maths.Points,
                'thinking_mark' : complete_thinking.Points,
                'writing_mark' : complete_writing.Points
            }
            self.agg_data = pd.DataFrame(self.agg_data)
            
            # compute selective mark
            self.agg_data['overall_mark'] = self.agg_data['reading_mark'] + self.agg_data['maths_mark'] + self.agg_data['thinking_mark'] + self.agg_data['writing_mark']

            # Rank info
            self.agg_data['reading_rank'] = complete_reading.ranking
            self.agg_data['maths_rank'] = complete_maths.ranking
            self.agg_data['thinking_rank'] = complete_thinking.ranking
            self.agg_data['writing_rank'] = complete_writing.ranking
            
            # Overall ranking
            overall_sorted = -np.sort(-self.agg_data.overall_mark)
            self.agg_data['overall_rank'] = [np.where(overall_sorted == mark)[0][0]+1 for mark in self.agg_data.overall_mark]
            
            # add on answer sheets
            self.agg_data = self.agg_data.join(complete_reading.iloc[:, -1-self.reading_total:-1]).join(complete_maths.iloc[:, -1-self.maths_total:-1]).join(complete_thinking.iloc[:, -1-self.thinking_total:-1])
            
            
            # Deal with incomplete students
            has_reading = self.reading_data.iloc[[not student in complete_students for student in self.reading_data.SID], ].reset_index(drop = True)
            has_maths = self.maths_data.iloc[[not student in complete_students for student in self.maths_data.SID], ].reset_index(drop = True)
            has_thinking = self.thinking_data.iloc[[not student in complete_students for student in self.thinking_data.SID], ].reset_index(drop = True)
            has_writing = self.writing_data.iloc[[not student in complete_students for student in self.writing_data.SID], ].reset_index(drop = True)
            
            # list of incomplete students
            incomplete_students = np.concatenate([has_reading.SID, has_maths.SID, has_thinking.SID, has_writing.SID])
            incomplete_students = np.unique(incomplete_students)
                        
            # incomplete student infos
            incomplete_df = self.generate_incomplete_df(incomplete_students, has_reading, has_maths, has_thinking, has_writing)

            
            # Combine incomplete dataframe with complete
            self.agg_data = pd.concat([self.agg_data, incomplete_df], ignore_index = True)
            self.agg_data.iloc[:, 12:] = self.agg_data.iloc[:, 12:].astype('float64')
            return(self.agg_data)

    
    
    # Generate complete pdf
    def complete_pdf(self, student_index, test_name):
        if self.test_type == 'octt':
            student_info = self.agg_data.iloc[student_index, ]
            name = student_info[0]
            eng_mark = student_info.reading_mark
            maths_mark = student_info.maths_mark
            thinking_mark = student_info.thinking_mark
            oc_mark = student_info.oc_mark
            eng_rank = student_info.reading_rank
            maths_rank = student_info.maths_rank
            thinking_rank = student_info.thinking_rank
            overall_rank = student_info.overall_rank
            
            
            # Calculate percentiles for each subject
            if eng_rank/self.num_reading > 0.5:
                eng_perc = 50
            elif eng_rank/self.num_reading > 0.25:
                eng_perc = 25
            elif eng_rank/self.num_reading > 0.1:
                eng_perc = 15
            else:
                eng_perc = 10
                
            if maths_rank/self.num_maths > 0.5:
                maths_perc = 50
            elif maths_rank/self.num_maths > 0.25:
                maths_perc = 25
            elif maths_rank/self.num_maths > 0.1:
                maths_perc = 15
            else:
                maths_perc = 10
            
            if thinking_rank/self.num_thinking > 0.5:
                thinking_perc = 50
            elif thinking_rank/self.num_thinking > 0.25:
                thinking_perc = 25
            elif thinking_rank/self.num_thinking > 0.1:
                thinking_perc = 15
            else:
                thinking_perc = 10
            
            # Initialise pdf and settings
            pdf = PDF()
            pdf.set_auto_page_break(auto = True)
            pdf.set_margin(15)
            pdf.add_font('cambria', '', 
                    'pdf_resources/Cambria.ttf',
                    uni = True)
            pdf.add_font('cambria', 'B', 
                    'pdf_resources/cambria-bold.ttf',
                    uni = True)
            pdf.add_font('cambria', 'I', 
                    'pdf_resources/cambria-italic.ttf',
                    uni = True)
            
            
            # Page 1: Summary Page
            pdf.add_page()
            
            # heading
            pdf.set_font('cambria', style = 'B', size = 23)
            pdf.cell(0, 10, test_name, 0, 1, 'C')
            pdf.set_font('cambria', style = '', size = 18)
            pdf.cell(0, 10, 'Student: {}'.format(name), 0, 1, 'C')
            pdf.ln(5)
            
            # Summary Text
            pdf.set_font('cambria', 'I', size = 14)
            pdf.multi_cell(0, 5, 
                           '''The following report contains cohort information on \
the exams that the student has completed. Please email the branch manager \
if there are any concerns or errors.''')
            pdf.ln(3)
            
            # Summary Table
            pdf.set_font('cambria', '', size = 14)
            summary_table = [['Summary', 'Mark'],
                             ['Reading', '{}/{}'.format(int(eng_mark), self.reading_total)],
                             ['Mathematical Reasoning', '{}/{}'.format(int(maths_mark), self.maths_total)],
                             ['Thinking Skills', '{}/{}'.format(int(thinking_mark), self.thinking_total)],
                             ['Overall Rank', '{}/{}'.format(int(overall_rank), int(self.num_complete))]]
            pdf.create_table(table_data = summary_table, data_size = 14, cell_width = [150, 30],
                             emphasize_data = ['Marks', 'Summary', 'Overall Rank'],
                             emphasize_style = 'B')

            pdf.ln(3)
            pdf.image('percentile_bands/reading' + str(eng_perc) + '.png', w = 200, x= 0.5)
            pdf.image('percentile_bands/maths' + str(maths_perc) + '.png', w = 200, x= 0.5)
            pdf.image('percentile_bands/thinking' + str(thinking_perc) + '.png', w = 200, x= 0.5)
            
            # Create caption
            pdf.ln(3)
            pdf.set_font('cambria', 'I', size = 11)
            pdf.multi_cell(0, 5, 
                           'Note: Highlighted band in light blue indicates the percentile band your child performed in.')
            
            
            
            pdf.add_page()
            pdf.set_font('cambria', style = '', size = 10)
            pdf.multi_cell(0, 5,
                           '''The following table shows all the OCTT schools ranked. The colour coding is an \
indication of the attainability given your child's performance in this exam. 

Red -> Difficult
Orange -> Attainable
Green -> Safe Option''')
            
            pdf.ln(3)
            pdf.set_font('cambria', style = 'B', size = 11)
            pdf.cell(0, 5, 'Student Percentile: {:.2f}%'.format(overall_rank/self.num_complete*100), 0, 0, 'C')
            pdf.ln(4.5)
            pdf.school_table_perc_oc(overall_rank/self.num_complete)
            
            
            
            
            
            # Page 2: Reading Comprehension
            pdf.add_page()
            pdf.new_section('Reading')
            pdf.ln(5)
            
            # Table of statistics
            eng_table = [['Student Mark:', '{}/{}'.format(int(eng_mark), self.reading_total)],
                         ['Student Rank:', '{}/{}'.format(int(eng_rank), self.num_reading)],
                         ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.reading_mark), self.reading_total)],
                         ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.reading_mark)), self.reading_total)]]
            pdf.set_font('cambria', '', 14)
            pdf.create_mark_table(eng_table)
            pdf.ln(5)
            
            # Create chart and put in image
            reading_fig = plots.reading_chart(self.agg_data, eng_mark)
            reading_fig.savefig('reading_fig.png', dpi = 500)
            plt.close(reading_fig)
            pdf.image('reading_fig.png', x = 30, w = 150)    
            
            # Table of Question Breakdown
            eng_q_table = [['Question Type', 'Questions', 'Mark'],
                           ['Normal Texts', 'Q1-11', '{}/11'.format(int(np.sum(student_info[10:21])))],
                           ['Cloze Passages', 'Q12-17', '{}/6'.format(int(np.sum(student_info[21:27])))],
                           ['Comparing Extracts', 'Q18-25', '{}/8'.format(int(np.sum(student_info[27:35])))]]
            pdf.create_table(eng_q_table, title = 'Question Breakdown', data_size = 14, title_size = 15, 
                             emphasize_data = ['Question Type', 'Questions', 'Mark'], emphasize_style = 'B',
                             cell_width = [100, 60, 20])
            

            
            # Page 3: All Reading Comprehension Questions
            pdf.add_page()
            pdf.set_font('cambria', 'B', size = 14)
            pdf.cell(10, txt = 'Reading Test Breakdown')
            pdf.ln(5)
            pdf.set_font('cambria', '', 10)
            pdf.cell(0, 5, txt = '''Questions highlighted green indicate that the student \
responded correctly. Red highlight indicates an incorrect response.''')
            pdf.ln(5)
            reading_qtable = pdf.questions_table('reading_q_types.xlsx', self.agg_data.iloc[:, 10:35])
            pdf.create_question_table(reading_qtable, self.agg_data.iloc[student_index, 10:35], did_test = True)
            
            
            
            
            # Page 4: Mathematics
            pdf.add_page()
            pdf.new_section('Mathematical Reasoning')
            pdf.ln(5)
            
            # Table of Statistics
            maths_table = [['Student Mark:', '{}/{}'.format(int(maths_mark), int(self.maths_total))],
                           ['Student Rank:', '{}/{}'.format(int(maths_rank), int(self.num_maths))],
                           ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.maths_mark), int(self.maths_total))],
                           ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.maths_mark)), int(self.maths_total))]]
            pdf.set_font('cambria', '', 14)
            pdf.create_mark_table(maths_table)
            pdf.ln(5)
            
            # Create chart and put in image
            maths_fig = plots.maths_chart(self.agg_data, maths_mark)
            maths_fig.savefig('maths_fig.png', dpi = 500)
            plt.close(maths_fig)
            pdf.image('maths_fig.png', x = 30, w = 150)
            
            # Message to say no maths question breakdown
            pdf.set_font('cambria', style = 'I', size = 14)
            pdf.multi_cell(0, 5,
                           txt = '''* Common questions breakdown unavailable for mathematical \
reasoning as there are only 1-3 questions per category.''')
            
            
            
            # Page 5: All maths questions
            pdf.add_page()
            pdf.set_font('cambria', 'B', size = 14)
            pdf.cell(10, txt = 'Mathematics Test Breakdown')
            pdf.ln(5)
            pdf.set_font('cambria', '', 10)
            pdf.cell(0, 5, txt = 'Questions highlighted green indicate that the student responded correctly. Red highlight indicates an incorrect response.')
            pdf.ln(5)
            maths_qtable = pdf.questions_table('maths_q_types.xlsx', self.agg_data.iloc[:, 35:70])
            pdf.create_question_table(maths_qtable, self.agg_data.iloc[student_index, 35:70], did_test = True)
            
            
            
            # Page 6: Thinking Skills
            pdf.add_page()
            pdf.new_section('Thinking Skills')
            pdf.ln(5)
            
            # table of statistics
            thinking_table = [['Student Mark:', '{}/{}'.format(int(thinking_mark), int(self.thinking_total))],
                              ['Student Rank:', '{}/{}'.format(int(thinking_rank), int(self.num_thinking))],
                              ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.thinking_mark), int(self.thinking_total))],
                              ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.thinking_mark)), int(self.thinking_total))]]
            pdf.set_font('cambria', '', 14)
            pdf.create_mark_table(thinking_table)
            pdf.ln(5)
            
            # create chart, and put in image
            thinking_fig = plots.thinking_chart(self.agg_data, thinking_mark)
            thinking_fig.savefig('thinking_fig.png', dpi = 500)
            plt.close(thinking_fig)
            pdf.image('thinking_fig.png', x = 30, w = 150)
            
            # Table of Common Questions student got wrong
            thinking_q = pd.read_excel('thinking_q_types.xlsx').iloc[:, 0]
            top_tq = Counter(thinking_q).most_common()[0:4]
            t_marks = []
            for q in top_tq:
                t_marks.append(np.sum(student_info.iloc[70:100].reset_index(drop = True)[thinking_q == q[0]]))
            
            t_common_q = [['Most Common Questions', 'Mark']]
            t_common_q2 = [[top_tq[i][0], '{}/{}'.format(int(t_marks[i]), top_tq[i][1])] for i in range(4)]
            t_common_q = t_common_q + t_common_q2
            
            pdf.create_table(t_common_q, title = 'Question Breakdown', data_size = 14, 
                             title_size = 15, emphasize_data = ['Most Common Questions', 'Mark'],
                             emphasize_style = 'B', cell_width = [150, 30])
            
            
            
            
            # Page 7: Every Thinking Skills Question
            pdf.add_page()
            pdf.set_font('cambria', 'B', size = 14)
            pdf.cell(10, txt = 'Thinking Skills Test Breakdown')
            pdf.ln(5)
            pdf.set_font('cambria', '', 10)
            pdf.cell(0, 5, txt = 'Questions highlighted green indicate that the student responded correctly. Red highlight indicates an incorrect response.')
            pdf.ln(5)
            thinking_qtable = pdf.questions_table('thinking_q_types.xlsx', self.agg_data.iloc[:, 70:100])
            pdf.create_question_table(thinking_qtable, self.agg_data.iloc[student_index, 70:100], did_test = True)
            
            
            
            
            # Output pdf
            pdf_name = '{}.pdf'.format(name)
            pdf.output(pdf_name, 'F')
            
            # Merge report pdf with solutions
            merged_pdf = PdfMerger()
            merged_pdf.append(pdf_name)
            merged_pdf.append('octt_sols.pdf')
            merged_pdf.write(pdf_name)
            merged_pdf.close()
        
        
        
        
        elif self.test_type == 'sttc':
            student_info = self.agg_data.iloc[student_index, ]
            name = student_info.iloc[0]
            eng_mark = student_info.reading_mark
            maths_mark = student_info.maths_mark
            thinking_mark = student_info.thinking_mark
            writing_mark = student_info.writing_mark
            eng_rank = student_info.reading_rank
            maths_rank = student_info.maths_rank
            thinking_rank = student_info.thinking_rank
            writing_rank = student_info.writing_rank
            overall_rank = student_info.overall_rank
            
            # Calculate percentiles for each subject
            if eng_rank/self.num_reading > 0.5:
                eng_perc = 50
            elif eng_rank/self.num_reading > 0.25:
                eng_perc = 25
            elif eng_rank/self.num_reading > 0.1:
                eng_perc = 15
            else:
                eng_perc = 10
                
            if maths_rank/self.num_maths > 0.5:
                maths_perc = 50
            elif maths_rank/self.num_maths > 0.25:
                maths_perc = 25
            elif maths_rank/self.num_maths > 0.1:
                maths_perc = 15
            else:
                maths_perc = 10
            
            if thinking_rank/self.num_thinking > 0.5:
                thinking_perc = 50
            elif thinking_rank/self.num_thinking > 0.25:
                thinking_perc = 25
            elif thinking_rank/self.num_thinking > 0.1:
                thinking_perc = 15
            else:
                thinking_perc = 10
            
            if writing_rank/self.num_writing > 0.5:
                writing_perc = 50
            elif writing_rank/self.num_writing > 0.25:
                writing_perc = 25
            elif writing_rank/self.num_writing > 0.1:
                writing_perc = 15
            else:
                writing_perc = 10
            
            
            # Initialise pdf and settings
            pdf = PDF()
            pdf.set_auto_page_break(auto = True)
            pdf.set_margin(15)
            pdf.add_font('cambria', '', 
                    'pdf_resources/Cambria.ttf',
                    uni = True)
            pdf.add_font('cambria', 'B', 
                    'pdf_resources/cambria-bold.ttf',
                    uni = True)
            pdf.add_font('cambria', 'I', 
                    'pdf_resources/cambria-italic.ttf',
                    uni = True)
            
            # Page 1: Summary Page
            pdf.add_page()
            pdf.set_font('cambria', style = 'B', size = 23)
            pdf.cell(0, 10, test_name, 0, 1, 'C')
            pdf.set_font('cambria', style = '', size = 18)
            pdf.cell(0, 10, 'Student: {}'.format(name), 0, 1, 'C')
            pdf.ln(5)
            
            
            # Summary text
            pdf.set_font('cambria', 'I', size = 14)
            pdf.multi_cell(0, 5, 
                           '''The following report contains cohort information on \
the exams that the student has completed. Please email the branch manager \
if there are any concerns or errors.''')
            pdf.ln(3)
            
            # Summary Text
            pdf.set_font('cambria', '', size = 14)
            summary_table = [['Summary', 'Mark'],
                             ['Reading', '{}/{}'.format(int(eng_mark), self.reading_total)],
                             ['Mathematical Reasoning','{}/{}'.format(int(maths_mark), self.maths_total)], 
                             ['Thinking Skills', '{}/{}'.format(int(thinking_mark), self.thinking_total)],
                             ['Writing', '{}/{}'.format(int(writing_mark), self.writing_total)],
                             ['Overall Rank', '{}/{}'.format(int(overall_rank), int(self.num_complete))]]
            pdf.create_table(table_data = summary_table, data_size = 14, cell_width = [150, 30],  
                            emphasize_data = ['Marks', 'Summary', 'Overall Rank'], 
                            emphasize_style = 'B')
            
            pdf.ln(3)
            pdf.image('percentile_bands/reading' + str(eng_perc) + '.png', w = 200, x= 0.5)
            pdf.image('percentile_bands/maths' + str(maths_perc) + '.png', w = 200, x= 0.5)
            pdf.image('percentile_bands/thinking' + str(thinking_perc) + '.png', w = 200, x= 0.5)
            pdf.image('percentile_bands/writing' + str(writing_perc) + '.png', w = 200, x= 0.5)
            
            # Create caption
            pdf.ln(3)
            pdf.set_font('cambria', 'I', size = 11)
            pdf.multi_cell(0, 5, 
                           'Note: Highlighted band in light blue indicates the percentile band your child performed in.')
            
            
            
            
            
            # Page 2: School Rankings page
            pdf.add_page()
            pdf.set_font('cambria', style = '', size = 10)
            pdf.multi_cell(0, 5,
                           '''The following table shows all the selective schools ranked. The colour coding is an \
indication of the attainability given your child's performance in this exam. 

Red -> Difficult
Orange -> Attainable
Green -> Safe Option

Schools denoted with a '(P)' are partial selective schools.''')
            
            pdf.ln(3)
            pdf.set_font('cambria', style = 'B', size = 11)
            pdf.cell(0, 5, 'Student Percentile: {:.2f}%'.format(overall_rank/self.num_complete*100), 0, 0, 'C')
            pdf.ln(4.5)
            pdf.school_table_perc(overall_rank/self.num_complete)
            
            
            
            # Page 3: Reading Comprehension
            pdf.add_page()
            pdf.new_section('Reading')
            pdf.ln(5)
            
            # table of statistics
            eng_table = [['Student Mark:', '{}/{}'.format(int(eng_mark), self.reading_total)],
                         ['Student Rank:', '{}/{}'.format(int(eng_rank), self.num_reading)],
                         ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.reading_mark), self.reading_total)],
                         ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.reading_mark)), self.reading_total)]]
            pdf.set_font('cambria', '', 14)
            pdf.create_mark_table(eng_table)
            
            pdf.ln(0.5)
           
            # Create chart and put in image
            eng_fig = plots.reading_chart(self.agg_data, eng_mark)
            eng_fig.savefig('eng_fig.png', dpi = 500)
            plt.close(eng_fig)
            pdf.image('eng_fig.png', x = 30, w = 150)
            
            # Table of Questions breakdown
            eng_q_table = [['Question Type', 'Questions', 'Mark'],
                           ['Normal Texts', 'Q1-14', '{}/14'.format(int(np.sum(student_info[12:26])))],
                           ['Cloze Passages', 'Q15-20', '{}/6'.format(int(np.sum(student_info[26:32])))],
                           ['Comparing Extracts', 'Q21-30', '{}/10'.format(int(np.sum(student_info[32:42])))]]
            pdf.create_table(eng_q_table, title = 'Question Breakdown', data_size = 14, title_size = 15, 
                             emphasize_data = ['Question Type', 'Questions', 'Mark'], emphasize_style = 'B',
                             cell_width = [100, 60, 20])
            
            
            
            
            # Page 4: All Reading Comprehension Questions
            pdf.add_page()
            pdf.set_font('cambria', 'B', size = 14)
            pdf.cell(10, txt = 'Reading Test Breakdown')
            pdf.ln(5)
            pdf.set_font('cambria', '', 10)
            pdf.cell(0, 5, txt = 'Questions highlighted green indicate that the student responded correctly. Red highlight indicates an incorrect response.')
            pdf.ln(5)
            reading_qtable = pdf.questions_table('reading_q_types.xlsx', self.agg_data.iloc[:, 12:42])
            pdf.create_question_table(reading_qtable, self.agg_data.iloc[student_index, 12:42], did_test = True)
            
            
            
            
            # Page 5: Mathematical Reasoning
            pdf.add_page()
            pdf.new_section('Mathematical Reasoning')
            pdf.ln(5)
            
            # Table of statistics
            maths_table = [['Student Mark:', '{}/{}'.format(int(maths_mark), int(self.maths_total))],
                           ['Student Rank:', '{}/{}'.format(int(maths_rank), int(self.num_maths))],
                           ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.maths_mark), int(self.maths_total))],
                           ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.maths_mark)), int(self.maths_total))]]
            pdf.set_font('cambria', '', 14)
            pdf.create_mark_table(maths_table)
            
            pdf.ln(0.5)
            
            # Create chart and put in image
            maths_fig = plots.maths_chart(self.agg_data, maths_mark)
            maths_fig.savefig('maths_fig.png', dpi = 500)
            plt.close(maths_fig)
            pdf.image('maths_fig.png', x = 30, w = 150)
            
            # message to say no maths question breakdown
            pdf.set_font('cambria', style = 'I', size = 14)
            pdf.multi_cell(0, 5, 
                           txt = '* Common questions breakdown unavailable for mathematical reasoning as there are only 1-3 questions per category.')
            
            
            
            
            # Page 6: Every Mathematical Reasoning Question
            pdf.add_page()
            pdf.set_font('cambria', 'B', size = 14)
            pdf.cell(10, txt = 'Mathematics Test Breakdown')
            pdf.ln(5)
            pdf.set_font('cambria', '', 10)
            pdf.cell(0, 5, txt = 'Questions highlighted green indicate that the student responded correctly. Red highlight indicates an incorrect response.')
            pdf.ln(5)
            maths_qtable = pdf.questions_table('maths_q_types.xlsx', self.agg_data.iloc[:, 42:77])
            pdf.create_question_table(maths_qtable, self.agg_data.iloc[student_index, 42:77], did_test = True)
            
            
            
            # Page 7: Thinking Skills
            pdf.add_page()
            pdf.new_section('Thinking Skills')
            pdf.ln(5)
            
            # Table of statistics
            thinking_table = [['Student Mark:', '{}/{}'.format(int(thinking_mark), int(self.thinking_total))],
                              ['Student Rank:', '{}/{}'.format(int(thinking_rank), int(self.num_thinking))],
                              ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.thinking_mark), int(self.thinking_total))],
                              ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.thinking_mark)), int(self.thinking_total))]]
            pdf.set_font('cambria', '', 14)
            pdf.create_mark_table(thinking_table)
            
            pdf.ln(0.5)
            
            # Create chart and put in image
            thinking_fig = plots.thinking_chart(self.agg_data, thinking_mark)
            thinking_fig.savefig('thinking_fig.png', dpi = 500)
            plt.close(thinking_fig)
            pdf.image('thinking_fig.png', x = 30, w = 150)
            
            # create table of common questions student got wrong
            thinking_q = pd.read_excel('thinking_q_types.xlsx').iloc[:, 0]
            top_tq = Counter(thinking_q).most_common()[0:4]
            t_marks = []
            for q in top_tq:
                t_marks.append(np.sum(student_info.iloc[77:117].reset_index(drop = True)[thinking_q == q[0]]))
            
            t_common_q = [['Most Common Questions', 'Mark']]
            t_common_q2 = [[top_tq[i][0], '{}/{}'.format(int(t_marks[i]), top_tq[i][1])] for i in range(0, 4)]
            t_common_q = t_common_q + t_common_q2
            
            pdf.create_table(t_common_q, title = 'Question Breakdown', data_size = 14,
                            title_size = 15, emphasize_data = ['Most Common Questions', 'Mark'], emphasize_style = 'B',
                            cell_width = [150, 30])
            
            
            
            
            # Page 8: Every Thinking Skills Question
            pdf.add_page()
            pdf.set_font('cambria', 'B', size = 14)
            pdf.cell(10, txt = 'Thinking Skills Test Breakdown')
            pdf.ln(5)
            pdf.set_font('cambria', '', 10)
            pdf.cell(0, 5, txt = 'Questions highlighted green indicate that the student responded correctly. Red highlight indicates an incorrect response.')
            pdf.ln(5)
            thinking_qtable = pdf.questions_table('thinking_q_types.xlsx', self.agg_data.iloc[:, 77:117])
            pdf.create_question_table(thinking_qtable, self.agg_data.iloc[student_index, 77:117], did_test = True)
                    
            
            
            
            # Page 9 : Writing
            pdf.add_page()
            pdf.new_section('Writing')
            pdf.ln(5)
            
            # table of statistics
            writing_table = [['Student Mark:', '{}/{}'.format(int(writing_mark), int(self.writing_total))],
                             ['Student Rank:', '{}/{}'.format(int(writing_rank), int(self.num_writing))],
                             ['Cohort Average:', '{:.2f}/{}'.format(np.mean(self.agg_data.writing_mark), int(self.writing_total))],
                             ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.writing_mark)), int(self.writing_total))]]
            pdf.set_font('cambria', '', 14)
            pdf.create_mark_table(writing_table)
            
            pdf.ln(0.5)
            
            # create chart, and put in image
            writing_fig = plots.writing_chart(self.agg_data, writing_mark)
            writing_fig.savefig('writing_fig.png', dpi = 500)
            plt.close(writing_fig)
            pdf.image('writing_fig.png', x = 30, w = 150)
            
            
            
            # Output pdf
            pdf_name =  '{}.pdf'.format(name)
            pdf.output(pdf_name, 'F')
          
            # create merged pdf with solutions
            merged_pdf = PdfMerger()
            merged_pdf.append(pdf_name)
            merged_pdf.append('sttc_sols.pdf')
            merged_pdf.write(pdf_name)
            merged_pdf.close()
        
        
        
        
        
        else: # wemt
            student_info = self.agg_data.iloc[student_index, ]
            name = student_info[0]
            eng_mark = student_info.reading_mark
            maths_mark = student_info.maths_mark
            thinking_mark = student_info.thinking_mark
            writing_mark = student_info.writing_mark
            eng_rank = student_info.reading_rank
            maths_rank = student_info.maths_rank
            thinking_rank = student_info.thinking_rank
            writing_rank = student_info.writing_rank
            overall_mark = student_info.overall_mark
            overall_rank = student_info.overall_rank

            # Initialise pdf and settings
            pdf = PDF()
            pdf.set_auto_page_break(auto = True)
            pdf.set_margin(15)
            pdf.add_font('cambria', '', 
                    'pdf_resources/Cambria.ttf',
                    uni = True)
            pdf.add_font('cambria', 'B', 
                    'pdf_resources/cambria-bold.ttf',
                    uni = True)
            pdf.add_font('cambria', 'I', 
                    'pdf_resources/cambria-italic.ttf',
                    uni = True)
            
            # Page 1: Summary Page
            pdf.add_page()
            pdf.set_font('cambria', style = 'B', size = 23)
            pdf.cell(0, 10, test_name, 0, 1, 'C')
            pdf.set_font('cambria', style = '', size = 18)
            pdf.cell(0, 10, 'Student: {}'.format(name), 0, 1, 'C')
            pdf.ln(5)
            
            
            # Summary text
            pdf.set_font('cambria', 'I', size = 14)
            pdf.multi_cell(0, 5, 
                           '''The following report contains cohort information on \
the exams that the student has completed. Please email the branch manager \
if there are any concerns or errors.''')
            pdf.ln(3)
            
            # Summary Text
            pdf.set_font('cambria', '', size = 14)
            summary_table = [['Summary', 'Mark'],
                             ['Reading', '{}/{}'.format(int(eng_mark), self.reading_total)],
                             ['Mathematical Reasoning','{}/{}'.format(int(maths_mark), self.maths_total)], 
                             ['Thinking Skills', '{}/{}'.format(int(thinking_mark), self.thinking_total)],
                             ['Writing', '{}/{}'.format(int(writing_mark), self.writing_total)],
                             ['Total Score', '{}/73'.format(int(overall_mark))],
                             ['Overall Rank', '{}/{}'.format(int(overall_rank), int(self.num_complete))]]
            pdf.create_table(table_data = summary_table, data_size = 14, cell_width = [150, 30],  
                            emphasize_data = ['Marks', 'Summary', 'Total Score', 'Overall Rank'], 
                            emphasize_style = 'B')
            
            
            
            
            
            # Page 2: Reading Comprehension
            pdf.add_page()
            pdf.new_section('Reading')
            pdf.ln(5)
            
            # table of statistics
            eng_table = [['Student Mark:', '{}/{}'.format(int(eng_mark), self.reading_total)],
                         ['Student Rank:', '{}/{}'.format(int(eng_rank), self.num_reading)],
                         ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.reading_mark), self.reading_total)],
                         ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.reading_mark)), self.reading_total)]]
            pdf.set_font('cambria', '', 14)
            pdf.create_mark_table(eng_table)
            
            pdf.ln(0.5)
           
            # Create chart and put in image
            eng_fig = plots.reading_chart(self.agg_data, eng_mark)
            eng_fig.savefig('eng_fig.png', dpi = 500)
            plt.close(eng_fig)
            pdf.image('eng_fig.png', x = 30, w = 150)
            
            # Table of Questions breakdown
            eng_q_table = [['Question Type', 'Questions', 'Mark'],
                           ['Normal Texts', 'Q1-14', '{}/14'.format(int(np.sum(student_info[12:17])))],
                           ['Cloze Passages', 'Q15-20', '{}/6'.format(int(np.sum(student_info[17:22])))],
                           ['Comparing Extracts', 'Q21-30', '{}/10'.format(int(np.sum(student_info[22:27])))]]
            pdf.create_table(eng_q_table, title = 'Question Breakdown', data_size = 14, title_size = 15, 
                             emphasize_data = ['Question Type', 'Questions', 'Mark'], emphasize_style = 'B',
                             cell_width = [100, 60, 20])
            
            
            
            
            # Page 3: All Reading Comprehension Questions
            pdf.add_page()
            pdf.set_font('cambria', 'B', size = 14)
            pdf.cell(10, txt = 'Reading Test Breakdown')
            pdf.ln(5)
            pdf.set_font('cambria', '', 10)
            pdf.cell(0, 5, txt = 'Questions highlighted green indicate that the student responded correctly. Red highlight indicates an incorrect response.')
            pdf.ln(5)
            reading_qtable = pdf.questions_table('reading_q_types.xlsx', self.agg_data.iloc[:, 12:27])
            pdf.create_question_table(reading_qtable, self.agg_data.iloc[student_index, 12:27], did_test = True)
            
            
            
            
            # Page 5: Mathematical Reasoning
            pdf.add_page()
            pdf.new_section('Mathematical Reasoning')
            pdf.ln(5)
            
            # Table of statistics
            maths_table = [['Student Mark:', '{}/{}'.format(int(maths_mark), int(self.maths_total))],
                           ['Student Rank:', '{}/{}'.format(int(maths_rank), int(self.num_maths))],
                           ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.maths_mark), int(self.maths_total))],
                           ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.maths_mark)), int(self.maths_total))]]
            pdf.set_font('cambria', '', 14)
            pdf.create_mark_table(maths_table)
            
            pdf.ln(0.5)
            
            # Create chart and put in image
            maths_fig = plots.maths_chart(self.agg_data, maths_mark)
            maths_fig.savefig('maths_fig.png', dpi = 500)
            plt.close(maths_fig)
            pdf.image('maths_fig.png', x = 30, w = 150)
            
            # message to say no maths question breakdown
            pdf.set_font('cambria', style = 'I', size = 14)
            pdf.multi_cell(0, 5, 
                           txt = '* Common questions breakdown unavailable for mathematical reasoning as there are only 1-3 questions per category.')
            
            
            
            
            # Page 6: Every Mathematical Reasoning Question
            pdf.add_page()
            pdf.set_font('cambria', 'B', size = 14)
            pdf.cell(10, txt = 'Mathematics Test Breakdown')
            pdf.ln(5)
            pdf.set_font('cambria', '', 10)
            pdf.cell(0, 5, txt = 'Questions highlighted green indicate that the student responded correctly. Red highlight indicates an incorrect response.')
            pdf.ln(5)
            maths_qtable = pdf.questions_table('maths_q_types.xlsx', self.agg_data.iloc[:, 27:45])
            pdf.create_question_table(maths_qtable, self.agg_data.iloc[student_index, 27:45], did_test = True)
            
            
            
            # Page 7: Thinking Skills
            pdf.add_page()
            pdf.new_section('Thinking Skills')
            pdf.ln(5)
            
            # Table of statistics
            thinking_table = [['Student Mark:', '{}/{}'.format(int(thinking_mark), int(self.thinking_total))],
                              ['Student Rank:', '{}/{}'.format(int(thinking_rank), int(self.num_thinking))],
                              ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.thinking_mark), int(self.thinking_total))],
                              ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.thinking_mark)), int(self.thinking_total))]]
            pdf.set_font('cambria', '', 14)
            pdf.create_mark_table(thinking_table)
            
            pdf.ln(0.5)
            
            # Create chart and put in image
            thinking_fig = plots.thinking_chart(self.agg_data, thinking_mark)
            thinking_fig.savefig('thinking_fig.png', dpi = 500)
            plt.close(thinking_fig)
            pdf.image('thinking_fig.png', x = 30, w = 150)
            
            # create table of common questions student got wrong
            thinking_q = pd.read_excel('thinking_q_types.xlsx').iloc[:, 0]
            top_tq = Counter(thinking_q).most_common()[0:4]
            t_marks = []
            for q in top_tq:
                t_marks.append(np.sum(student_info.iloc[45:65].reset_index(drop = True)[thinking_q == q[0]]))
            
            t_common_q = [['Most Common Questions', 'Mark']]
            t_common_q2 = [[top_tq[i][0], '{}/{}'.format(int(t_marks[i]), top_tq[i][1])] for i in range(0, 4)]
            t_common_q = t_common_q + t_common_q2
            
            pdf.create_table(t_common_q, title = 'Question Breakdown', data_size = 14,
                            title_size = 15, emphasize_data = ['Most Common Questions', 'Mark'], emphasize_style = 'B',
                            cell_width = [150, 30])
            
            
            
            
            # Page 8: Every Thinking Skills Question
            pdf.add_page()
            pdf.set_font('cambria', 'B', size = 14)
            pdf.cell(10, txt = 'Thinking Skills Test Breakdown')
            pdf.ln(5)
            pdf.set_font('cambria', '', 10)
            pdf.cell(0, 5, txt = 'Questions highlighted green indicate that the student responded correctly. Red highlight indicates an incorrect response.')
            pdf.ln(5)
            thinking_qtable = pdf.questions_table('thinking_q_types.xlsx', self.agg_data.iloc[:, 45:65])
            pdf.create_question_table(thinking_qtable, self.agg_data.iloc[student_index, 45:65], did_test = True)
                    
            
            
            
            # Page 9 : Writing
            pdf.add_page()
            pdf.new_section('Writing')
            pdf.ln(5)
            
            # table of statistics
            writing_table = [['Student Mark:', '{}/{}'.format(int(writing_mark), int(self.writing_total))],
                             ['Student Rank:', '{}/{}'.format(int(writing_rank), int(self.num_writing))],
                             ['Cohort Average:', '{:.2f}/{}'.format(np.mean(self.agg_data.writing_mark), int(self.writing_total))],
                             ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.writing_mark)), int(self.writing_total))]]
            pdf.set_font('cambria', '', 14)
            pdf.create_mark_table(writing_table)
            
            pdf.ln(0.5)
            
            # create chart, and put in image
            writing_fig = plots.writing_chart(self.agg_data, writing_mark)
            writing_fig.savefig('writing_fig.png', dpi = 500)
            plt.close(writing_fig)
            pdf.image('writing_fig.png', x = 30, w = 150)
            
            # output pdf
            pdf_name =  '{}.pdf'.format(name)
            pdf.output(pdf_name, 'F')
          
            # create merged pdf with solutions
            merged_pdf = PdfMerger()
            merged_pdf.append(pdf_name)
            merged_pdf.append('wemt_sols.pdf')
            merged_pdf.write(pdf_name)
            merged_pdf.close()
          
            
            
                
    # Generate incomplete pdf
    def incomplete_pdf(self, student_index, test_name):
        if self.test_type == 'octt':
            student_info = self.agg_data.iloc[student_index, ]
            name = student_info[0]
            did_eng = False
            did_maths = False
            did_thinking = False
            if ~np.isnan(student_info.reading_mark):
                eng_mark = int(student_info.reading_mark)
                eng_rank = int(student_info.reading_rank)
                
                # calculate percentile
                if eng_rank/self.num_reading > 0.5:
                    eng_perc = 50
                elif eng_rank/self.num_reading > 0.25:
                    eng_perc = 25
                elif eng_rank/self.num_reading > 0.1:
                    eng_perc = 15
                else:
                    eng_perc = 10
                
                did_eng = True
            if ~np.isnan(student_info.maths_mark):
                maths_mark = int(student_info.maths_mark)
                maths_rank = int(student_info.maths_rank)
                
                if maths_rank/self.num_maths > 0.5:
                    maths_perc = 50
                elif maths_rank/self.num_maths > 0.25:
                    maths_perc = 25
                elif maths_rank/self.num_maths > 0.1:
                    maths_perc = 15
                else:
                    maths_perc = 10
                
                did_maths = True
            if ~np.isnan(student_info.thinking_mark):
                thinking_mark = int(student_info.thinking_mark)
                thinking_rank = int(student_info.thinking_rank)
                
                if thinking_rank/self.num_thinking > 0.5:
                    thinking_perc = 50
                elif thinking_rank/self.num_thinking > 0.25:
                    thinking_perc = 25
                elif thinking_rank/self.num_thinking > 0.1:
                    thinking_perc = 15
                else:
                    thinking_perc = 10
                
                did_thinking = True
            
            # Pdf creation:
            # initialise pdf and settings
            pdf = PDF()
            pdf.set_auto_page_break(auto = True)
            pdf.set_margin(15)
            pdf.add_font('cambria', '', 
                    'pdf_resources/Cambria.ttf',
                    uni = True)
            pdf.add_font('cambria', 'B', 
                    'pdf_resources/cambria-bold.ttf',
                    uni = True)
            pdf.add_font('cambria', 'I', 
                    'pdf_resources/cambria-italic.ttf',
                    uni = True)
            
            
            
            
            # Page 1: Summary Page
            pdf.add_page()
            pdf.set_font('cambria', style = 'B', size = 23)
            pdf.cell(0, 10, test_name, 0, 1, 'C')
            pdf.set_font('cambria', '', size = 18)
            pdf.cell(0, 10, 'Student: {}'.format(name), 0, 1, 'C')
            pdf.ln(5)
            
            pdf.set_font('cambria', 'I', size = 14)
            pdf.multi_cell(0, 5, 
                        'The following report contains information on the {} ' \
                            'exams that the student has completed. As the student ' \
                                'did not complete every exam, the report is incomplete, including only ' \
                                    'data on the exams taken.'.format(np.sum((did_eng, did_maths, did_thinking))))
            
            pdf.ln(3)
            
            pdf.set_font('cambria', '', size = 14)
            summary_table = [['Summary', 'Mark']]
            if did_eng:
                summary_table.append(['Reading', '{}/{}'.format(eng_mark, int(self.reading_total))])
            else:
                summary_table.append(['Reading', 'NA'])
            if did_maths:
                summary_table.append(['Mathematical Reasoning', '{}/{}'.format(maths_mark, int(self.maths_total))])
            else:
                summary_table.append(['Mathematical Reasoning', 'NA'])
            if did_thinking:
                summary_table.append(['Thinking Skills', '{}/{}'.format(thinking_mark, int(self.thinking_total))])
            else:
                summary_table.append(['Thinking Skills', 'NA'])
            
            # put in NA's for OC mark and rank
            summary_table.append(['Overall Rank', 'NA'])
            
            pdf.create_table(table_data = summary_table, data_size = 14, cell_width = [150, 30],  
                            emphasize_data = ['Marks', 'Summary'], 
                            emphasize_style = 'B')
            
            pdf.ln(3)
            if did_eng:
                pdf.image('percentile_bands/reading' + str(eng_perc) + '.png', w = 200, x= 0.5)
            if did_maths:
                pdf.image('percentile_bands/maths' + str(maths_perc) + '.png', w = 200, x= 0.5)
            if did_thinking:
                pdf.image('percentile_bands/thinking' + str(thinking_perc) + '.png', w = 200, x= 0.5)
            
            
            
            # Page 2 & 3: Reading Comprehension
            if did_eng:
                pdf.add_page()
                pdf.new_section('Reading')
                pdf.ln(5)
            
                # table of statistics
                eng_table = [['Student Mark:', '{}/{}'.format(int(eng_mark), self.reading_total)],
                            ['Student Rank:', '{}/{}'.format(int(eng_rank), self.num_reading)],
                            ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.reading_mark), self.reading_total)],
                            ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.reading_mark)), self.reading_total)]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(eng_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                eng_fig = plots.reading_chart(self.agg_data, eng_mark)
                eng_fig.savefig('eng_fig.png', dpi = 500)
                plt.close(eng_fig)
                pdf.image('eng_fig.png', x = 30, w = 150)
                
                # table of question breakdown
                eng_q_table = [['Question Type', 'Questions', 'Mark'],
                               ['Normal Texts', 'Q1-11', '{}/11'.format(int(np.sum(student_info[10:21])))],
                               ['Cloze Passages', 'Q12-17', '{}/6'.format(int(np.sum(student_info[21:27])))],
                               ['Comparing Extracts', 'Q18-25', '{}/8'.format(int(np.sum(student_info[27:35])))]]
                pdf.create_table(eng_q_table, title = 'Question Breakdown', data_size = 14, title_size = 15, 
                                emphasize_data = ['Question Type', 'Questions', 'Mark'], emphasize_style = 'B',
                                cell_width = [100, 60, 20])
                
                
                pdf.add_page()
                pdf.set_font('cambria', 'B', size = 14)
                pdf.cell(10, txt = 'Reading Test Breakdown')
                pdf.ln(5)
                pdf.set_font('cambria', '', 10)
                pdf.cell(0, 5, txt = 'Questions highlighted green indicate that the student responded correctly. Red highlight indicates an incorrect response.')
                pdf.ln(5)
                reading_qtable = pdf.questions_table('reading_q_types.xlsx', self.agg_data.iloc[:, 10:35])
                pdf.create_question_table(reading_qtable, ans_sheet = self.agg_data.iloc[student_index, 10:35], did_test = True)
                
                
            else:
                pdf.add_page()
                pdf.new_section('Reading')
                pdf.ln(5)
            
                # table of statistics
                eng_table = [['Student Mark:', 'NA'],
                            ['Student Rank:', 'NA'],
                            ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.reading_mark), self.reading_total)],
                            ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.reading_mark)), self.reading_total)]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(eng_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                eng_fig = plots.reading_chart(self.agg_data, -1)
                eng_fig.savefig('eng_fig.png', dpi = 500)
                plt.close(eng_fig)
                pdf.image('eng_fig.png', x = 30, w = 150)
                
                
                pdf.add_page()
                pdf.set_font('cambria', 'B', size = 14)
                pdf.cell(10, txt = 'Reading Test Breakdown')
                pdf.ln(5)
                reading_qtable = pdf.questions_table('reading_q_types.xlsx', self.agg_data.iloc[:, 10:35])
                pdf.create_question_table(reading_qtable, ans_sheet = self.agg_data.iloc[student_index, 10:35], did_test = False)
                
            
            
            
            
            
            # Page 4 & 5: Mathematical Reasoning
            if did_maths:
                # Page 3 : Mathematics
                pdf.add_page()
                pdf.new_section('Mathematical Reasoning')
                pdf.ln(5)
                
                # table of statistics
                maths_table = [['Student Mark:', '{}/{}'.format(int(maths_mark), int(self.maths_total))],
                               ['Student Rank:', '{}/{}'.format(int(maths_rank), int(self.num_maths))],
                               ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.maths_mark), int(self.maths_total))],
                               ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.maths_mark)), int(self.num_maths))]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(maths_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                maths_fig = plots.maths_chart(self.agg_data, maths_mark)
                maths_fig.savefig('maths_fig.png', dpi = 500)
                plt.close(maths_fig)
                pdf.image('maths_fig.png', x = 30, w = 150)
                
                pdf.set_font('cambria', style = 'I', size = 14)
                pdf.multi_cell(0, 5, 
                            txt = '* Common questions breakdown unavailable for mathematical reasoning as there are only 1-3 questions per category.')
                    
            
                # new page for every Maths Question
                pdf.add_page()
                pdf.set_font('cambria', 'B', size = 14)
                pdf.cell(10, txt = 'Mathematics Test Breakdown')
                pdf.ln(5)
                pdf.set_font('cambria', '', 10)
                pdf.cell(0, 5, txt = 'Questions highlighted green indicate that the student responded correctly. Red highlight indicates an incorrect response.')
                pdf.ln(5)
                maths_qtable = pdf.questions_table('maths_q_types.xlsx', self.agg_data.iloc[:, 35:70])
                pdf.create_question_table(maths_qtable, self.agg_data.iloc[student_index, 35:70], did_test = True)
                
            else:
                pdf.add_page()
                pdf.new_section('Mathematical Reasoning')
                pdf.ln(5)
                
                # table of statistics
                maths_table = [['Student Mark:', 'NA'],
                            ['Student Rank:', 'NA'],
                            ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.maths_mark), int(self.maths_total))],
                            ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.maths_mark)), int(self.maths_total))]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(maths_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                maths_fig = plots.maths_chart(self.agg_data, -1)
                maths_fig.savefig('maths_fig.png', dpi = 500)
                plt.close(maths_fig)
                pdf.image('maths_fig.png', x = 30, w = 150)
                
                
                # new page for every Maths Question
                pdf.add_page()
                pdf.set_font('cambria', 'B', size = 14)
                pdf.cell(10, txt = 'Mathematics Test Breakdown')
                pdf.ln(5)
                maths_qtable = pdf.questions_table('maths_q_types.xlsx', self.agg_data.iloc[:, 35:70])
                pdf.create_question_table(maths_qtable, self.agg_data.iloc[student_index, 35:70], did_test = False)
            
            
            
            
            
            
            # Page 6 & 7: Thinking Skills
            if did_thinking:
                pdf.add_page()
                pdf.new_section('Thinking Skills')
                pdf.ln(5)
                
                # table of statistics
                thinking_table = [['Student Mark:', '{}/{}'.format(int(thinking_mark), int(self.thinking_total))],
                                ['Student Rank:', '{}/{}'.format(int(thinking_rank), int(self.num_thinking))],
                                ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.thinking_mark), int(self.thinking_total))],
                                ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.thinking_mark)), int(self.thinking_total))]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(thinking_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                thinking_fig = plots.thinking_chart(self.agg_data, thinking_mark)
                thinking_fig.savefig('thinking_fig.png', dpi = 500)
                plt.close(thinking_fig)
                pdf.image('thinking_fig.png', x = 30, w = 150)
                
                
                # create table of common questions student got wrong
                thinking_q = pd.read_excel('thinking_q_types.xlsx').q_type
                top_tq = Counter(thinking_q).most_common()[0:4]
                t_marks = []
                for q in top_tq:
                    t_marks.append(np.sum(student_info.iloc[70:100].reset_index(drop = True)[thinking_q == q[0]]))
                
                t_common_q = [['Most Common Questions', 'Mark']]
                t_common_q2 = [[top_tq[i][0], '{}/{}'.format(int(t_marks[i]), top_tq[i][1])] for i in range(0, 4)]
                t_common_q = t_common_q + t_common_q2
                
                pdf.create_table(t_common_q, title = 'Question Breakdown', data_size = 14,
                                title_size = 15, emphasize_data = ['Most Common Questions', 'Mark'], emphasize_style = 'B',
                                cell_width = [150, 30])
                
                
                pdf.add_page()
                pdf.set_font('cambria', 'B', size = 14)
                pdf.cell(10, txt = 'Thinking Skills Test Breakdown')
                pdf.ln(5)
                pdf.set_font('cambria', '', 10)
                pdf.cell(0, 5, txt = 'Questions highlighted green indicate that the student responded correctly. Red highlight indicates an incorrect response.')
                pdf.ln(5)
                thinking_qtable = pdf.questions_table('thinking_q_types.xlsx', self.agg_data.iloc[:, 70:100])
                pdf.create_question_table(thinking_qtable, self.agg_data.iloc[student_index, 70:100], did_test = True)
            
            else:
                pdf.add_page()
                pdf.new_section('Thinking Skills')
                pdf.ln(5)
                
                # table of statistics
                thinking_table = [['Student Mark:', 'NA'],
                                ['Student Rank:', 'NA'],
                                ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.thinking_mark), int(self.thinking_total))],
                                ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.thinking_mark)), int(self.thinking_total))]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(thinking_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                thinking_fig = plots.thinking_chart(self.agg_data, -1)
                thinking_fig.savefig('thinking_fig.png', dpi = 500)
                plt.close(thinking_fig)
                pdf.image('thinking_fig.png', x = 30, w = 150)
                
                
                # Page 6 : Every Thinking Skills Question
                pdf.add_page()
                pdf.set_font('cambria', 'B', size = 14)
                pdf.cell(10, txt = 'Thinking Skills Test Breakdown')
                pdf.ln(5)
                thinking_qtable = pdf.questions_table('thinking_q_types.xlsx', self.agg_data.iloc[:, 70:100])
                pdf.create_question_table(thinking_qtable, self.agg_data.iloc[student_index, 70:100], did_test = False)
                
            
            # Output pdf
            pdf_name =  '{}.pdf'.format(name)
            pdf.output(pdf_name, 'F')

            # create merged pdf with solutions
            merged_pdf = PdfMerger()
            merged_pdf.append(pdf_name)
            merged_pdf.append('octt_sols.pdf')
            merged_pdf.write(pdf_name)
            merged_pdf.close()
            
        
        elif self.test_type == 'sttc':
            student_info = self.agg_data.iloc[student_index, ]
            name = student_info[0]
            did_eng = False
            did_maths = False
            did_thinking = False
            did_writing = False
            if ~np.isnan(student_info.reading_mark):
                eng_mark = int(student_info.reading_mark)
                eng_rank = int(student_info.reading_rank)
                
                # calculate percentile
                if eng_rank/self.num_reading > 0.5:
                    eng_perc = 50
                elif eng_rank/self.num_reading > 0.25:
                    eng_perc = 25
                elif eng_rank/self.num_reading > 0.1:
                    eng_perc = 15
                else:
                    eng_perc = 10
                
                did_eng = True
            if ~np.isnan(student_info.maths_mark):
                maths_mark = int(student_info.maths_mark)
                maths_rank = int(student_info.maths_rank)
                
                if maths_rank/self.num_maths > 0.5:
                    maths_perc = 50
                elif maths_rank/self.num_maths > 0.25:
                    maths_perc = 25
                elif maths_rank/self.num_maths > 0.1:
                    maths_perc = 15
                else:
                    maths_perc = 10
                
                did_maths = True
            if ~np.isnan(student_info.thinking_mark):
                thinking_mark = int(student_info.thinking_mark)
                thinking_rank = int(student_info.thinking_rank)
                
                if thinking_rank/self.num_thinking > 0.5:
                    thinking_perc = 50
                elif thinking_rank/self.num_thinking > 0.25:
                    thinking_perc = 25
                elif thinking_rank/self.num_thinking > 0.1:
                    thinking_perc = 15
                else:
                    thinking_perc = 10
                
                did_thinking = True
                
            if ~np.isnan(student_info.writing_mark):
                writing_mark = int(student_info.writing_mark)
                writing_rank = int(student_info.writing_rank)
                
                if writing_rank/self.num_writing > 0.5:
                    writing_perc = 50
                elif writing_rank/self.num_writing  > 0.25:
                    writing_perc = 25
                elif writing_rank/self.num_writing  > 0.1:
                    writing_perc = 15
                else:
                    writing_perc = 10
                
                did_writing = True
            
            # Pdf creation:
            # initialise pdf and settings
            pdf = PDF()
            pdf.set_auto_page_break(auto = True)
            pdf.set_margin(15)
            pdf.add_font('cambria', '', 
                    'pdf_resources/Cambria.ttf',
                    uni = True)
            pdf.add_font('cambria', 'B', 
                    'pdf_resources/cambria-bold.ttf',
                    uni = True)
            pdf.add_font('cambria', 'I', 
                    'pdf_resources/cambria-italic.ttf',
                    uni = True)
            
            
            
            # Page 1: Summary page
            pdf.add_page()
            pdf.set_font('cambria', style = 'B', size = 23)
            pdf.cell(0, 10, test_name, 0, 1, 'C')
            pdf.set_font('cambria', '', size = 18)
            pdf.cell(0, 10, 'Student: {}'.format(name), 0, 1, 'C')
            pdf.ln(5)
            
            pdf.set_font('cambria', 'I', size = 14)
            pdf.multi_cell(0, 5, 
                        'The following report contains information on the {} ' \
                            'exams that the student has completed. As the student ' \
                                'did not complete every exam, the report is incomplete, including only ' \
                                    'data on the exams taken.'.format(np.sum((did_eng, did_maths, did_thinking, did_writing))))
            
            pdf.ln(3)
            
            pdf.set_font('cambria', '', size = 14)
            summary_table = [['Summary', 'Mark']]
            if did_eng:
                summary_table.append(['Reading', '{}/{}'.format(eng_mark, int(self.reading_total))])
            else:
                summary_table.append(['Reading', 'NA'])
            if did_maths:
                summary_table.append(['Mathematical Reasoning', '{}/{}'.format(maths_mark, int(self.maths_total))])
            else:
                summary_table.append(['Mathematical Reasoning', 'NA'])
            if did_thinking:
                summary_table.append(['Thinking Skills', '{}/{}'.format(thinking_mark, int(self.thinking_total))])
            else:
                summary_table.append(['Thinking Skills', 'NA'])
            if did_writing: 
                summary_table.append(['Writing', '{}/{}'.format(writing_mark, int(self.writing_total))])
            else:
                summary_table.append(['Writing', 'NA'])
            
            summary_table.append(['Overall Rank', 'NA'])
            
            pdf.create_table(table_data = summary_table, data_size = 14, cell_width = [150, 30],  
                            emphasize_data = ['Marks', 'Summary'], 
                            emphasize_style = 'B')
            
            pdf.ln(3)
            if did_eng:
                pdf.image('percentile_bands/reading' + str(eng_perc) + '.png', w = 200, x= 0.5)
            if did_maths:
                pdf.image('percentile_bands/maths' + str(maths_perc) + '.png', w = 200, x= 0.5)
            if did_thinking:
                pdf.image('percentile_bands/thinking' + str(thinking_perc) + '.png', w = 200, x= 0.5)
            if did_writing:
                pdf.image('percentile_bands/writing' + str(writing_perc) + '.png', w = 200, x= 0.5)
            
            
            
            # Page 2 & 3: Reading Comprehension
            if did_eng:
                pdf.add_page()
                pdf.new_section('Reading')
                pdf.ln(5)
            
                # table of statistics
                eng_table = [['Student Mark:', '{}/{}'.format(int(eng_mark), self.reading_total)],
                            ['Student Rank:', '{}/{}'.format(int(eng_rank), self.num_reading)],
                            ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.reading_mark), self.reading_total)],
                            ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.reading_mark)), self.reading_total)]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(eng_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                eng_fig = plots.reading_chart(self.agg_data, eng_mark)
                eng_fig.savefig('eng_fig.png', dpi = 500)
                plt.close(eng_fig)
                pdf.image('eng_fig.png', x = 30, w = 150)
                
                # table of question breakdown
                eng_q_table = [['Question Type', 'Questions', 'Mark'],
                            ['Normal Texts', 'Q1-14', '{}/14'.format(int(np.sum(student_info[12:26])))],
                            ['Cloze Passages', 'Q15-20', '{}/6'.format(int(np.sum(student_info[26:32])))],
                            ['Comparing Extracts', 'Q21-30', '{}/10'.format(int(np.sum(student_info[32:42])))]]
                pdf.create_table(eng_q_table, title = 'Question Breakdown', data_size = 14, title_size = 15, 
                                emphasize_data = ['Question Type', 'Questions', 'Mark'], emphasize_style = 'B',
                                cell_width = [100, 60, 20])
                
                # every reading comp question
                pdf.add_page()
                pdf.set_font('cambria', 'B', size = 14)
                pdf.cell(10, txt = 'Reading Test Breakdown')
                pdf.ln(5)
                pdf.set_font('cambria', '', 10)
                pdf.cell(0, 5, txt = 'Questions highlighted green indicate that the student responded correctly. Red highlight indicates an incorrect response.')
                pdf.ln(5)
                reading_qtable = pdf.questions_table('reading_q_types.xlsx', self.agg_data.iloc[:, 12:42])
                pdf.create_question_table(reading_qtable, self.agg_data.iloc[student_index, 12:42], did_test = True)
            
                
            else:
                pdf.add_page()
                pdf.new_section('Reading')
                pdf.ln(5)
            
                # table of statistics
                eng_table = [['Student Mark:', 'NA'],
                            ['Student Rank:', 'NA'],
                            ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.reading_mark), self.reading_total)],
                            ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.reading_mark)), self.reading_total)]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(eng_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                eng_fig = plots.reading_chart(self.agg_data, -1)
                eng_fig.savefig('eng_fig.png', dpi = 500)
                plt.close(eng_fig)
                pdf.image('eng_fig.png', x = 30, w = 150)
            
                # every reading comp question
                pdf.add_page()
                pdf.set_font('cambria', 'B', size = 14)
                pdf.cell(10, txt = 'Reading Test Breakdown')
                pdf.ln(5)
                reading_qtable = pdf.questions_table('reading_q_types.xlsx', self.agg_data.iloc[:, 12:42])
                pdf.create_question_table(reading_qtable, self.agg_data.iloc[student_index, 12:42], did_test = False)
            
            
            # Page 4 & 5: Mathematical Reasoning
            if did_maths:
                pdf.add_page()
                pdf.new_section('Mathematical Reasoning')
                pdf.ln(5)
                
                # table of statistics
                maths_table = [['Student Mark:', '{}/{}'.format(int(maths_mark), int(self.maths_total))],
                            ['Student Rank:', '{}/{}'.format(int(maths_rank), int(self.num_maths))],
                            ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.maths_mark), int(self.maths_total))],
                            ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.maths_mark)), int(self.maths_total))]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(maths_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                maths_fig = plots.maths_chart(self.agg_data, maths_mark)
                maths_fig.savefig('maths_fig.png', dpi = 500)
                plt.close(maths_fig)
                pdf.image('maths_fig.png', x = 30, w = 150)
                
                # message to say no maths question breakdown
                pdf.set_font('cambria', style = 'I', size = 14)
                pdf.multi_cell(0, 5, 
                            txt = '* The maths question breakdown will come at a later date while we revamp the question classification system.')
                
                
                
                # new page for every Maths Question
                pdf.add_page()
                pdf.set_font('cambria', 'B', size = 14)
                pdf.cell(10, txt = 'Mathematics Test Breakdown')
                pdf.ln(5)
                pdf.set_font('cambria', '', 10)
                pdf.cell(0, 5, txt = 'Questions highlighted green indicate that the student responded correctly. Red highlight indicates an incorrect response.')
                pdf.ln(5)
                maths_qtable = pdf.questions_table('maths_q_types.xlsx', self.agg_data.iloc[:, 42:77])
                pdf.create_question_table(maths_qtable, self.agg_data.iloc[student_index, 42:77], did_test = True)
                
            else:
                # Page 3 : Mathematics
                pdf.add_page()
                pdf.new_section('Mathematical Reasoning')
                pdf.ln(5)
                
                # table of statistics
                maths_table = [['Student Mark:', 'NA'],
                            ['Student Rank:', 'NA'],
                            ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.maths_mark), int(self.maths_total))],
                            ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.maths_mark)), int(self.maths_total))]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(maths_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                maths_fig = plots.maths_chart(self.agg_data, -1)
                maths_fig.savefig('maths_fig.png', dpi = 500)
                plt.close(maths_fig)
                pdf.image('maths_fig.png', x = 30, w = 150)
                
                
                # new page for every Maths Question
                pdf.add_page()
                pdf.set_font('cambria', 'B', size = 14)
                pdf.cell(10, txt = 'Mathematics Test Breakdown')
                pdf.ln(5)
                maths_qtable = pdf.questions_table('maths_q_types.xlsx', self.agg_data.iloc[:, 42:77])
                pdf.create_question_table(maths_qtable, self.agg_data.iloc[student_index, 42:77], did_test = False)
            
            # Page 6 & 7: Thinking Skills
            if did_thinking:
                pdf.add_page()
                pdf.new_section('Thinking Skills')
                pdf.ln(5)
                
                # table of statistics
                thinking_table = [['Student Mark:', '{}/{}'.format(int(thinking_mark), int(self.thinking_total))],
                                ['Student Rank:', '{}/{}'.format(int(thinking_rank), int(self.num_thinking))],
                                ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.thinking_mark), int(self.thinking_total))],
                                ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.thinking_mark)), int(self.thinking_total))]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(thinking_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                thinking_fig = plots.thinking_chart(self.agg_data, thinking_mark)
                thinking_fig.savefig('thinking_fig.png', dpi = 500)
                plt.close(thinking_fig)
                pdf.image('thinking_fig.png', x = 30, w = 150)
                
                
                # create table of common questions student got wrong
                thinking_q = pd.read_excel('thinking_q_types.xlsx').q_type
                top_tq = Counter(thinking_q).most_common()[0:4]
                t_marks = []
                for q in top_tq:
                    t_marks.append(np.sum(student_info.iloc[77:117].reset_index(drop = True)[thinking_q == q[0]]))
                
                t_common_q = [['Most Common Questions', 'Mark']]
                t_common_q2 = [[top_tq[i][0], '{}/{}'.format(int(t_marks[i]), top_tq[i][1])] for i in range(0, 4)]
                t_common_q = t_common_q + t_common_q2
                
                pdf.create_table(t_common_q, title = 'Question Breakdown', data_size = 14,
                                title_size = 15, emphasize_data = ['Most Common Questions', 'Mark'], emphasize_style = 'B',
                                cell_width = [150, 30])
                
                
                pdf.add_page()
                pdf.set_font('cambria', 'B', size = 14)
                pdf.cell(10, txt = 'Thinking Skills Test Breakdown')
                pdf.ln(5)
                pdf.set_font('cambria', '', 10)
                pdf.cell(0, 5, txt = 'Questions highlighted green indicate that the student responded correctly. Red highlight indicates an incorrect response.')
                pdf.ln(5)
                thinking_qtable = pdf.questions_table('thinking_q_types.xlsx', self.agg_data.iloc[:, 77:117])
                pdf.create_question_table(thinking_qtable, self.agg_data.iloc[student_index, 77:117], did_test = True)
            
            else:
                pdf.add_page()
                pdf.new_section('Thinking Skills')
                pdf.ln(5)
                
                # table of statistics
                thinking_table = [['Student Mark:', 'NA'],
                                ['Student Rank:', 'NA'],
                                ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.thinking_mark), int(self.thinking_total))],
                                ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.thinking_mark)), int(self.thinking_total))]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(thinking_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                thinking_fig = plots.thinking_chart(self.agg_data, -1)
                thinking_fig.savefig('thinking_fig.png', dpi = 500)
                plt.close(thinking_fig)
                pdf.image('thinking_fig.png', x = 30, w = 150)
                
                # Page 6 : Every Thinking Skills Question
                pdf.add_page()
                pdf.set_font('cambria', 'B', size = 14)
                pdf.cell(10, txt = 'Thinking Skills Test Breakdown')
                pdf.ln(5)
                thinking_qtable = pdf.questions_table('thinking_q_types.xlsx', self.agg_data.iloc[:, 77:117])
                pdf.create_question_table(thinking_qtable, self.agg_data.iloc[student_index, 77:117], did_test = False)
            
            # create writing page if necessary
            if did_writing:
                pdf.add_page()
                pdf.new_section('Writing')
                pdf.ln(5)
                
                # table of statistics
                writing_table = [['Student Mark:', '{}/{}'.format(int(writing_mark), int(self.writing_total))],
                                ['Student Rank:', '{}/{}'.format(int(writing_rank), int(self.num_writing))],
                                ['Cohort Average:', '{:.2f}/{}'.format(np.mean(self.agg_data.writing_mark), self.writing_total)],
                                ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.writing_mark)), self.writing_total)]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(writing_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                writing_fig = plots.writing_chart(self.agg_data, writing_mark)
                writing_fig.savefig('writing_fig.png', dpi = 500)
                plt.close(writing_fig)
                pdf.image('writing_fig.png', x = 30, w = 150)
                
            else:
                pdf.add_page()
                pdf.new_section('Writing')
                pdf.ln(5)
                
                # table of statistics
                writing_table = [['Student Mark:', 'NA'],
                                ['Student Rank:', 'NA'],
                                ['Cohort Average:', '{:.2f}/{}'.format(np.mean(self.agg_data.writing_mark), self.writing_total)],
                                ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.writing_mark)), self.writing_total)]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(writing_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                writing_fig = plots.writing_chart(self.agg_data, -1)
                writing_fig.savefig('writing_fig.png', dpi = 500)
                plt.close(writing_fig)
                pdf.image('writing_fig.png', x = 30, w = 150)
                
            # output pdf
            pdf_name = '{}.pdf'.format(name)
            pdf.output(pdf_name, 'F')
            
            # create merged pdf with solutions
            merged_pdf = PdfMerger()
            merged_pdf.append(pdf_name)
            merged_pdf.append('sttc_sols.pdf')
            merged_pdf.write(pdf_name)
            merged_pdf.close()
        
        
        
        
        else: # wemt
            student_info = self.agg_data.iloc[student_index, ]
            name = student_info[0]
            did_eng = False
            did_maths = False
            did_thinking = False
            did_writing = False
            if ~np.isnan(student_info.reading_mark):
                eng_mark = int(student_info.reading_mark)
                eng_rank = int(student_info.reading_rank)
                did_eng = True
            if ~np.isnan(student_info.maths_mark):
                maths_mark = int(student_info.maths_mark)
                maths_rank = int(student_info.maths_rank)
                did_maths = True
            if ~np.isnan(student_info.thinking_mark):
                thinking_mark = int(student_info.thinking_mark)
                thinking_rank = int(student_info.thinking_rank)
                did_thinking = True
                
            if ~np.isnan(student_info.writing_mark):
                writing_mark = int(student_info.writing_mark)
                writing_rank = int(student_info.writing_rank)
                did_writing = True
            
            # Pdf creation:
            # initialise pdf and settings
            pdf = PDF()
            pdf.set_auto_page_break(auto = True)
            pdf.set_margin(15)
            pdf.add_font('cambria', '', 
                    'pdf_resources/Cambria.ttf',
                    uni = True)
            pdf.add_font('cambria', 'B', 
                    'pdf_resources/cambria-bold.ttf',
                    uni = True)
            pdf.add_font('cambria', 'I', 
                    'pdf_resources/cambria-italic.ttf',
                    uni = True)
            
            
            
            # Page 1: Summary page
            pdf.add_page()
            pdf.set_font('cambria', style = 'B', size = 23)
            pdf.cell(0, 10, test_name, 0, 1, 'C')
            pdf.set_font('cambria', '', size = 18)
            pdf.cell(0, 10, 'Student: {}'.format(name), 0, 1, 'C')
            pdf.ln(5)
            
            pdf.set_font('cambria', 'I', size = 14)
            pdf.multi_cell(0, 5, 
                        'The following report contains information on the {} ' \
                            'exams that the student has completed. As the student ' \
                                'did not complete every exam, the report is incomplete, including only ' \
                                    'data on the exams taken.'.format(np.sum((did_eng, did_maths, did_thinking, did_writing))))
            
            pdf.ln(3)
            
            pdf.set_font('cambria', '', size = 14)
            summary_table = [['Summary', 'Mark']]
            if did_eng:
                summary_table.append(['Reading', '{}/{}'.format(eng_mark, int(self.reading_total))])
            else:
                summary_table.append(['Reading', 'NA'])
            if did_maths:
                summary_table.append(['Mathematical Reasoning', '{}/{}'.format(maths_mark, int(self.maths_total))])
            else:
                summary_table.append(['Mathematical Reasoning', 'NA'])
            if did_thinking:
                summary_table.append(['Thinking Skills', '{}/{}'.format(thinking_mark, int(self.thinking_total))])
            else:
                summary_table.append(['Thinking Skills', 'NA'])
            if did_writing: 
                summary_table.append(['Writing', '{}/{}'.format(writing_mark, int(self.writing_total))])
            else:
                summary_table.append(['Writing', 'NA'])
            
            summary_table.append(['Overall Score', 'NA'])
            summary_table.append(['Overall Rank', 'NA'])
            
            pdf.create_table(table_data = summary_table, data_size = 14, cell_width = [150, 30],  
                            emphasize_data = ['Marks', 'Summary'], 
                            emphasize_style = 'B')
            
            
            
            # Page 2 & 3: Reading Comprehension
            if did_eng:
                pdf.add_page()
                pdf.new_section('Reading')
                pdf.ln(5)
            
                # table of statistics
                eng_table = [['Student Mark:', '{}/{}'.format(int(eng_mark), self.reading_total)],
                            ['Student Rank:', '{}/{}'.format(int(eng_rank), self.num_reading)],
                            ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.reading_mark), self.reading_total)],
                            ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.reading_mark)), self.reading_total)]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(eng_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                eng_fig = plots.reading_chart(self.agg_data, eng_mark)
                eng_fig.savefig('eng_fig.png', dpi = 500)
                plt.close(eng_fig)
                pdf.image('eng_fig.png', x = 30, w = 150)
                
                # table of question breakdown
                eng_q_table = [['Question Type', 'Questions', 'Mark'],
                            ['Normal Texts', 'Q1-14', '{}/14'.format(int(np.sum(student_info[12:17])))],
                            ['Cloze Passages', 'Q15-20', '{}/6'.format(int(np.sum(student_info[17:22])))],
                            ['Comparing Extracts', 'Q21-30', '{}/10'.format(int(np.sum(student_info[22:27])))]]
                pdf.create_table(eng_q_table, title = 'Question Breakdown', data_size = 14, title_size = 15, 
                                emphasize_data = ['Question Type', 'Questions', 'Mark'], emphasize_style = 'B',
                                cell_width = [100, 60, 20])
                
                # every reading comp question
                pdf.add_page()
                pdf.set_font('cambria', 'B', size = 14)
                pdf.cell(10, txt = 'Reading Test Breakdown')
                pdf.ln(5)
                pdf.set_font('cambria', '', 10)
                pdf.cell(0, 5, txt = 'Questions highlighted green indicate that the student responded correctly. Red highlight indicates an incorrect response.')
                pdf.ln(5)
                reading_qtable = pdf.questions_table('reading_q_types.xlsx', self.agg_data.iloc[:, 12:27])
                pdf.create_question_table(reading_qtable, self.agg_data.iloc[student_index, 12:27], did_test = True)
            
                
            else:
                pdf.add_page()
                pdf.new_section('Reading')
                pdf.ln(5)
            
                # table of statistics
                eng_table = [['Student Mark:', 'NA'],
                            ['Student Rank:', 'NA'],
                            ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.reading_mark), self.reading_total)],
                            ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.reading_mark)), self.reading_total)]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(eng_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                eng_fig = plots.reading_chart(self.agg_data, -1)
                eng_fig.savefig('eng_fig.png', dpi = 500)
                plt.close(eng_fig)
                pdf.image('eng_fig.png', x = 30, w = 150)
            
                # every reading comp question
                pdf.add_page()
                pdf.set_font('cambria', 'B', size = 14)
                pdf.cell(10, txt = 'Reading Test Breakdown')
                pdf.ln(5)
                reading_qtable = pdf.questions_table('reading_q_types.xlsx', self.agg_data.iloc[:, 12:27])
                pdf.create_question_table(reading_qtable, self.agg_data.iloc[student_index, 12:27], did_test = False)
            
            
            # Page 4 & 5: Mathematical Reasoning
            if did_maths:
                pdf.add_page()
                pdf.new_section('Mathematical Reasoning')
                pdf.ln(5)
                
                # table of statistics
                maths_table = [['Student Mark:', '{}/{}'.format(int(maths_mark), int(self.maths_total))],
                            ['Student Rank:', '{}/{}'.format(int(maths_rank), int(self.num_maths))],
                            ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.maths_mark), int(self.maths_total))],
                            ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.maths_mark)), int(self.maths_total))]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(maths_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                maths_fig = plots.maths_chart(self.agg_data, maths_mark)
                maths_fig.savefig('maths_fig.png', dpi = 500)
                plt.close(maths_fig)
                pdf.image('maths_fig.png', x = 30, w = 150)
                
                # message to say no maths question breakdown
                pdf.set_font('cambria', style = 'I', size = 14)
                pdf.multi_cell(0, 5, 
                            txt = '* The maths question breakdown will come at a later date while we revamp the question classification system.')
                
                
                
                # new page for every Maths Question
                pdf.add_page()
                pdf.set_font('cambria', 'B', size = 14)
                pdf.cell(10, txt = 'Mathematics Test Breakdown')
                pdf.ln(5)
                pdf.set_font('cambria', '', 10)
                pdf.cell(0, 5, txt = 'Questions highlighted green indicate that the student responded correctly. Red highlight indicates an incorrect response.')
                pdf.ln(5)
                maths_qtable = pdf.questions_table('maths_q_types.xlsx', self.agg_data.iloc[:, 27:45])
                pdf.create_question_table(maths_qtable, self.agg_data.iloc[student_index, 27:45], did_test = True)
                
            else:
                # Page 3 : Mathematics
                pdf.add_page()
                pdf.new_section('Mathematical Reasoning')
                pdf.ln(5)
                
                # table of statistics
                maths_table = [['Student Mark:', 'NA'],
                            ['Student Rank:', 'NA'],
                            ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.maths_mark), int(self.maths_total))],
                            ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.maths_mark)), int(self.maths_total))]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(maths_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                maths_fig = plots.maths_chart(self.agg_data, -1)
                maths_fig.savefig('maths_fig.png', dpi = 500)
                plt.close(maths_fig)
                pdf.image('maths_fig.png', x = 30, w = 150)
                
                
                # new page for every Maths Question
                pdf.add_page()
                pdf.set_font('cambria', 'B', size = 14)
                pdf.cell(10, txt = 'Mathematics Test Breakdown')
                pdf.ln(5)
                maths_qtable = pdf.questions_table('maths_q_types.xlsx', self.agg_data.iloc[:, 27:45])
                pdf.create_question_table(maths_qtable, self.agg_data.iloc[student_index, 27:45], did_test = False)
            
            # Page 6 & 7: Thinking Skills
            if did_thinking:
                pdf.add_page()
                pdf.new_section('Thinking Skills')
                pdf.ln(5)
                
                # table of statistics
                thinking_table = [['Student Mark:', '{}/{}'.format(int(thinking_mark), int(self.thinking_total))],
                                ['Student Rank:', '{}/{}'.format(int(thinking_rank), int(self.num_thinking))],
                                ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.thinking_mark), int(self.thinking_total))],
                                ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.thinking_mark)), int(self.thinking_total))]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(thinking_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                thinking_fig = plots.thinking_chart(self.agg_data, thinking_mark)
                thinking_fig.savefig('thinking_fig.png', dpi = 500)
                plt.close(thinking_fig)
                pdf.image('thinking_fig.png', x = 30, w = 150)
                
                
                # create table of common questions student got wrong
                thinking_q = pd.read_excel('thinking_q_types.xlsx').q_type
                top_tq = Counter(thinking_q).most_common()[0:4]
                t_marks = []
                for q in top_tq:
                    t_marks.append(np.sum(student_info.iloc[45:65].reset_index(drop = True)[thinking_q == q[0]]))
                
                t_common_q = [['Most Common Questions', 'Mark']]
                t_common_q2 = [[top_tq[i][0], '{}/{}'.format(int(t_marks[i]), top_tq[i][1])] for i in range(0, 4)]
                t_common_q = t_common_q + t_common_q2
                
                pdf.create_table(t_common_q, title = 'Question Breakdown', data_size = 14,
                                title_size = 15, emphasize_data = ['Most Common Questions', 'Mark'], emphasize_style = 'B',
                                cell_width = [150, 30])
                
                
                pdf.add_page()
                pdf.set_font('cambria', 'B', size = 14)
                pdf.cell(10, txt = 'Thinking Skills Test Breakdown')
                pdf.ln(5)
                pdf.set_font('cambria', '', 10)
                pdf.cell(0, 5, txt = 'Questions highlighted green indicate that the student responded correctly. Red highlight indicates an incorrect response.')
                pdf.ln(5)
                thinking_qtable = pdf.questions_table('thinking_q_types.xlsx', self.agg_data.iloc[:, 45:65])
                pdf.create_question_table(thinking_qtable, self.agg_data.iloc[student_index, 45:65], did_test = True)
            
            else:
                pdf.add_page()
                pdf.new_section('Thinking Skills')
                pdf.ln(5)
                
                # table of statistics
                thinking_table = [['Student Mark:', 'NA'],
                                ['Student Rank:', 'NA'],
                                ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.thinking_mark), int(self.thinking_total))],
                                ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.thinking_mark)), int(self.thinking_total))]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(thinking_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                thinking_fig = plots.thinking_chart(self.agg_data, -1)
                thinking_fig.savefig('thinking_fig.png', dpi = 500)
                plt.close(thinking_fig)
                pdf.image('thinking_fig.png', x = 30, w = 150)
                
                # Page 6 : Every Thinking Skills Question
                pdf.add_page()
                pdf.set_font('cambria', 'B', size = 14)
                pdf.cell(10, txt = 'Thinking Skills Test Breakdown')
                pdf.ln(5)
                thinking_qtable = pdf.questions_table('thinking_q_types.xlsx', self.agg_data.iloc[:, 45:65])
                pdf.create_question_table(thinking_qtable, self.agg_data.iloc[student_index, 45:65], did_test = False)
            
            # create writing page if necessary
            if did_writing:
                pdf.add_page()
                pdf.new_section('Writing')
                pdf.ln(5)
                
                # table of statistics
                writing_table = [['Student Mark:', '{}/{}'.format(int(writing_mark), int(self.writing_total))],
                                ['Student Rank:', '{}/{}'.format(int(writing_rank), int(self.num_writing))],
                                ['Cohort Average:', '{:.2f}/{}'.format(np.mean(self.agg_data.writing_mark), self.writing_total)],
                                ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.writing_mark)), self.writing_total)]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(writing_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                writing_fig = plots.writing_chart(self.agg_data, writing_mark)
                writing_fig.savefig('writing_fig.png', dpi = 500)
                plt.close(writing_fig)
                pdf.image('writing_fig.png', x = 30, w = 150)
                
            else:
                pdf.add_page()
                pdf.new_section('Writing')
                pdf.ln(5)
                
                # table of statistics
                writing_table = [['Student Mark:', 'NA'],
                                ['Student Rank:', 'NA'],
                                ['Cohort Average:', '{:.2f}/{}'.format(np.mean(self.agg_data.writing_mark), self.writing_total)],
                                ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.writing_mark)), self.writing_total)]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(writing_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                writing_fig = plots.writing_chart(self.agg_data, -1)
                writing_fig.savefig('writing_fig.png', dpi = 500)
                plt.close(writing_fig)
                pdf.image('writing_fig.png', x = 30, w = 150)
                
            # output pdf
            pdf_name = '{}.pdf'.format(name)
            pdf.output(pdf_name, 'F')
            
            # create merged pdf with solutions
            merged_pdf = PdfMerger()
            merged_pdf.append(pdf_name)
            merged_pdf.append('wemt_sols.pdf')
            merged_pdf.write(pdf_name)
            merged_pdf.close()
            
            
            
                
    # Generate incomplete pdf
    def incomplete_pdf(self, student_index, test_name):
        if self.test_type == 'octt':
            student_info = self.agg_data.iloc[student_index, ]
            name = student_info[0]
            did_eng = False
            did_maths = False
            did_thinking = False
            if ~np.isnan(student_info.reading_mark):
                eng_mark = int(student_info.reading_mark)
                eng_rank = int(student_info.reading_rank)
                
                # calculate percentile
                if eng_rank/self.num_reading > 0.5:
                    eng_perc = 50
                elif eng_rank/self.num_reading > 0.25:
                    eng_perc = 25
                elif eng_rank/self.num_reading > 0.1:
                    eng_perc = 15
                else:
                    eng_perc = 10
                
                did_eng = True
            if ~np.isnan(student_info.maths_mark):
                maths_mark = int(student_info.maths_mark)
                maths_rank = int(student_info.maths_rank)
                
                if maths_rank/self.num_maths > 0.5:
                    maths_perc = 50
                elif maths_rank/self.num_maths > 0.25:
                    maths_perc = 25
                elif maths_rank/self.num_maths > 0.1:
                    maths_perc = 15
                else:
                    maths_perc = 10
                
                did_maths = True
            if ~np.isnan(student_info.thinking_mark):
                thinking_mark = int(student_info.thinking_mark)
                thinking_rank = int(student_info.thinking_rank)
                
                if thinking_rank/self.num_thinking > 0.5:
                    thinking_perc = 50
                elif thinking_rank/self.num_thinking > 0.25:
                    thinking_perc = 25
                elif thinking_rank/self.num_thinking > 0.1:
                    thinking_perc = 15
                else:
                    thinking_perc = 10
                
                did_thinking = True
            
            # Pdf creation:
            # initialise pdf and settings
            pdf = PDF()
            pdf.set_auto_page_break(auto = True)
            pdf.set_margin(15)
            pdf.add_font('cambria', '', 
                    'pdf_resources/Cambria.ttf',
                    uni = True)
            pdf.add_font('cambria', 'B', 
                    'pdf_resources/cambria-bold.ttf',
                    uni = True)
            pdf.add_font('cambria', 'I', 
                    'pdf_resources/cambria-italic.ttf',
                    uni = True)
            
            
            
            
            # Page 1: Summary Page
            pdf.add_page()
            pdf.set_font('cambria', style = 'B', size = 23)
            pdf.cell(0, 10, test_name, 0, 1, 'C')
            pdf.set_font('cambria', '', size = 18)
            pdf.cell(0, 10, 'Student: {}'.format(name), 0, 1, 'C')
            pdf.ln(5)
            
            pdf.set_font('cambria', 'I', size = 14)
            pdf.multi_cell(0, 5, 
                        'The following report contains information on the {} ' \
                            'exams that the student has completed. As the student ' \
                                'did not complete every exam, the report is incomplete, including only ' \
                                    'data on the exams taken.'.format(np.sum((did_eng, did_maths, did_thinking))))
            
            pdf.ln(3)
            
            pdf.set_font('cambria', '', size = 14)
            summary_table = [['Summary', 'Mark']]
            if did_eng:
                summary_table.append(['Reading', '{}/{}'.format(eng_mark, int(self.reading_total))])
            else:
                summary_table.append(['Reading', 'NA'])
            if did_maths:
                summary_table.append(['Mathematical Reasoning', '{}/{}'.format(maths_mark, int(self.maths_total))])
            else:
                summary_table.append(['Mathematical Reasoning', 'NA'])
            if did_thinking:
                summary_table.append(['Thinking Skills', '{}/{}'.format(thinking_mark, int(self.thinking_total))])
            else:
                summary_table.append(['Thinking Skills', 'NA'])
            
            # put in NA's for OC mark and rank
            summary_table.append(['Overall Rank', 'NA'])
            
            pdf.create_table(table_data = summary_table, data_size = 14, cell_width = [150, 30],  
                            emphasize_data = ['Marks', 'Summary'], 
                            emphasize_style = 'B')
            
            pdf.ln(3)
            if did_eng:
                pdf.image('percentile_bands/reading' + str(eng_perc) + '.png', w = 200, x= 0.5)
            if did_maths:
                pdf.image('percentile_bands/maths' + str(maths_perc) + '.png', w = 200, x= 0.5)
            if did_thinking:
                pdf.image('percentile_bands/thinking' + str(thinking_perc) + '.png', w = 200, x= 0.5)
            
            
            
            # Page 2 & 3: Reading Comprehension
            if did_eng:
                pdf.add_page()
                pdf.new_section('Reading')
                pdf.ln(5)
            
                # table of statistics
                eng_table = [['Student Mark:', '{}/{}'.format(int(eng_mark), self.reading_total)],
                            ['Student Rank:', '{}/{}'.format(int(eng_rank), self.num_reading)],
                            ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.reading_mark), self.reading_total)],
                            ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.reading_mark)), self.reading_total)]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(eng_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                eng_fig = plots.reading_chart(self.agg_data, eng_mark)
                eng_fig.savefig('eng_fig.png', dpi = 500)
                plt.close(eng_fig)
                pdf.image('eng_fig.png', x = 30, w = 150)
                
                # table of question breakdown
                eng_q_table = [['Question Type', 'Questions', 'Mark'],
                               ['Normal Texts', 'Q1-11', '{}/11'.format(int(np.sum(student_info[10:21])))],
                               ['Cloze Passages', 'Q12-17', '{}/6'.format(int(np.sum(student_info[21:27])))],
                               ['Comparing Extracts', 'Q18-25', '{}/8'.format(int(np.sum(student_info[27:35])))]]
                pdf.create_table(eng_q_table, title = 'Question Breakdown', data_size = 14, title_size = 15, 
                                emphasize_data = ['Question Type', 'Questions', 'Mark'], emphasize_style = 'B',
                                cell_width = [100, 60, 20])
                
                
                pdf.add_page()
                pdf.set_font('cambria', 'B', size = 14)
                pdf.cell(10, txt = 'Reading Test Breakdown')
                pdf.ln(5)
                pdf.set_font('cambria', '', 10)
                pdf.cell(0, 5, txt = 'Questions highlighted green indicate that the student responded correctly. Red highlight indicates an incorrect response.')
                pdf.ln(5)
                reading_qtable = pdf.questions_table('reading_q_types.xlsx', self.agg_data.iloc[:, 10:35])
                pdf.create_question_table(reading_qtable, ans_sheet = self.agg_data.iloc[student_index, 10:35], did_test = True)
                
                
            else:
                pdf.add_page()
                pdf.new_section('Reading')
                pdf.ln(5)
            
                # table of statistics
                eng_table = [['Student Mark:', 'NA'],
                            ['Student Rank:', 'NA'],
                            ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.reading_mark), self.reading_total)],
                            ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.reading_mark)), self.reading_total)]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(eng_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                eng_fig = plots.reading_chart(self.agg_data, -1)
                eng_fig.savefig('eng_fig.png', dpi = 500)
                plt.close(eng_fig)
                pdf.image('eng_fig.png', x = 30, w = 150)
                
                
                pdf.add_page()
                pdf.set_font('cambria', 'B', size = 14)
                pdf.cell(10, txt = 'Reading Test Breakdown')
                pdf.ln(5)
                reading_qtable = pdf.questions_table('reading_q_types.xlsx', self.agg_data.iloc[:, 10:35])
                pdf.create_question_table(reading_qtable, ans_sheet = self.agg_data.iloc[student_index, 10:35], did_test = False)
                
            
            
            
            
            
            # Page 4 & 5: Mathematical Reasoning
            if did_maths:
                # Page 3 : Mathematics
                pdf.add_page()
                pdf.new_section('Mathematical Reasoning')
                pdf.ln(5)
                
                # table of statistics
                maths_table = [['Student Mark:', '{}/{}'.format(int(maths_mark), int(self.maths_total))],
                               ['Student Rank:', '{}/{}'.format(int(maths_rank), int(self.num_maths))],
                               ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.maths_mark), int(self.maths_total))],
                               ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.maths_mark)), int(self.num_maths))]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(maths_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                maths_fig = plots.maths_chart(self.agg_data, maths_mark)
                maths_fig.savefig('maths_fig.png', dpi = 500)
                plt.close(maths_fig)
                pdf.image('maths_fig.png', x = 30, w = 150)
                
                pdf.set_font('cambria', style = 'I', size = 14)
                pdf.multi_cell(0, 5, 
                            txt = '* Common questions breakdown unavailable for mathematical reasoning as there are only 1-3 questions per category.')
                    
            
                # new page for every Maths Question
                pdf.add_page()
                pdf.set_font('cambria', 'B', size = 14)
                pdf.cell(10, txt = 'Mathematics Test Breakdown')
                pdf.ln(5)
                pdf.set_font('cambria', '', 10)
                pdf.cell(0, 5, txt = 'Questions highlighted green indicate that the student responded correctly. Red highlight indicates an incorrect response.')
                pdf.ln(5)
                maths_qtable = pdf.questions_table('maths_q_types.xlsx', self.agg_data.iloc[:, 35:70])
                pdf.create_question_table(maths_qtable, self.agg_data.iloc[student_index, 35:70], did_test = True)
                
            else:
                pdf.add_page()
                pdf.new_section('Mathematical Reasoning')
                pdf.ln(5)
                
                # table of statistics
                maths_table = [['Student Mark:', 'NA'],
                            ['Student Rank:', 'NA'],
                            ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.maths_mark), int(self.maths_total))],
                            ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.maths_mark)), int(self.maths_total))]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(maths_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                maths_fig = plots.maths_chart(self.agg_data, -1)
                maths_fig.savefig('maths_fig.png', dpi = 500)
                plt.close(maths_fig)
                pdf.image('maths_fig.png', x = 30, w = 150)
                
                
                # new page for every Maths Question
                pdf.add_page()
                pdf.set_font('cambria', 'B', size = 14)
                pdf.cell(10, txt = 'Mathematics Test Breakdown')
                pdf.ln(5)
                maths_qtable = pdf.questions_table('maths_q_types.xlsx', self.agg_data.iloc[:, 35:70])
                pdf.create_question_table(maths_qtable, self.agg_data.iloc[student_index, 35:70], did_test = False)
            
            
            
            
            
            
            # Page 6 & 7: Thinking Skills
            if did_thinking:
                pdf.add_page()
                pdf.new_section('Thinking Skills')
                pdf.ln(5)
                
                # table of statistics
                thinking_table = [['Student Mark:', '{}/{}'.format(int(thinking_mark), int(self.thinking_total))],
                                ['Student Rank:', '{}/{}'.format(int(thinking_rank), int(self.num_thinking))],
                                ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.thinking_mark), int(self.thinking_total))],
                                ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.thinking_mark)), int(self.thinking_total))]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(thinking_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                thinking_fig = plots.thinking_chart(self.agg_data, thinking_mark)
                thinking_fig.savefig('thinking_fig.png', dpi = 500)
                plt.close(thinking_fig)
                pdf.image('thinking_fig.png', x = 30, w = 150)
                
                
                # create table of common questions student got wrong
                thinking_q = pd.read_excel('thinking_q_types.xlsx').q_type
                top_tq = Counter(thinking_q).most_common()[0:4]
                t_marks = []
                for q in top_tq:
                    t_marks.append(np.sum(student_info.iloc[70:100].reset_index(drop = True)[thinking_q == q[0]]))
                
                t_common_q = [['Most Common Questions', 'Mark']]
                t_common_q2 = [[top_tq[i][0], '{}/{}'.format(int(t_marks[i]), top_tq[i][1])] for i in range(0, 4)]
                t_common_q = t_common_q + t_common_q2
                
                pdf.create_table(t_common_q, title = 'Question Breakdown', data_size = 14,
                                title_size = 15, emphasize_data = ['Most Common Questions', 'Mark'], emphasize_style = 'B',
                                cell_width = [150, 30])
                
                
                pdf.add_page()
                pdf.set_font('cambria', 'B', size = 14)
                pdf.cell(10, txt = 'Thinking Skills Test Breakdown')
                pdf.ln(5)
                pdf.set_font('cambria', '', 10)
                pdf.cell(0, 5, txt = 'Questions highlighted green indicate that the student responded correctly. Red highlight indicates an incorrect response.')
                pdf.ln(5)
                thinking_qtable = pdf.questions_table('thinking_q_types.xlsx', self.agg_data.iloc[:, 70:100])
                pdf.create_question_table(thinking_qtable, self.agg_data.iloc[student_index, 70:100], did_test = True)
            
            else:
                pdf.add_page()
                pdf.new_section('Thinking Skills')
                pdf.ln(5)
                
                # table of statistics
                thinking_table = [['Student Mark:', 'NA'],
                                ['Student Rank:', 'NA'],
                                ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.thinking_mark), int(self.thinking_total))],
                                ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.thinking_mark)), int(self.thinking_total))]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(thinking_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                thinking_fig = plots.thinking_chart(self.agg_data, -1)
                thinking_fig.savefig('thinking_fig.png', dpi = 500)
                plt.close(thinking_fig)
                pdf.image('thinking_fig.png', x = 30, w = 150)
                
                
                # Page 6 : Every Thinking Skills Question
                pdf.add_page()
                pdf.set_font('cambria', 'B', size = 14)
                pdf.cell(10, txt = 'Thinking Skills Test Breakdown')
                pdf.ln(5)
                thinking_qtable = pdf.questions_table('thinking_q_types.xlsx', self.agg_data.iloc[:, 70:100])
                pdf.create_question_table(thinking_qtable, self.agg_data.iloc[student_index, 70:100], did_test = False)
                
            
            # Output pdf
            pdf_name =  '{}.pdf'.format(name)
            pdf.output(pdf_name, 'F')

            # create merged pdf with solutions
            merged_pdf = PdfMerger()
            merged_pdf.append(pdf_name)
            merged_pdf.append('octt_sols.pdf')
            merged_pdf.write(pdf_name)
            merged_pdf.close()
            
        
        elif self.test_type == 'sttc':
            student_info = self.agg_data.iloc[student_index, ]
            name = student_info[0]
            did_eng = False
            did_maths = False
            did_thinking = False
            did_writing = False
            if ~np.isnan(student_info.reading_mark):
                eng_mark = int(student_info.reading_mark)
                eng_rank = int(student_info.reading_rank)
                
                # calculate percentile
                if eng_rank/self.num_reading > 0.5:
                    eng_perc = 50
                elif eng_rank/self.num_reading > 0.25:
                    eng_perc = 25
                elif eng_rank/self.num_reading > 0.1:
                    eng_perc = 15
                else:
                    eng_perc = 10
                
                did_eng = True
            if ~np.isnan(student_info.maths_mark):
                maths_mark = int(student_info.maths_mark)
                maths_rank = int(student_info.maths_rank)
                
                if maths_rank/self.num_maths > 0.5:
                    maths_perc = 50
                elif maths_rank/self.num_maths > 0.25:
                    maths_perc = 25
                elif maths_rank/self.num_maths > 0.1:
                    maths_perc = 15
                else:
                    maths_perc = 10
                
                did_maths = True
            if ~np.isnan(student_info.thinking_mark):
                thinking_mark = int(student_info.thinking_mark)
                thinking_rank = int(student_info.thinking_rank)
                
                if thinking_rank/self.num_thinking > 0.5:
                    thinking_perc = 50
                elif thinking_rank/self.num_thinking > 0.25:
                    thinking_perc = 25
                elif thinking_rank/self.num_thinking > 0.1:
                    thinking_perc = 15
                else:
                    thinking_perc = 10
                
                did_thinking = True
                
            if ~np.isnan(student_info.writing_mark):
                writing_mark = int(student_info.writing_mark)
                writing_rank = int(student_info.writing_rank)
                
                if writing_rank/self.num_writing > 0.5:
                    writing_perc = 50
                elif writing_rank/self.num_writing  > 0.25:
                    writing_perc = 25
                elif writing_rank/self.num_writing  > 0.1:
                    writing_perc = 15
                else:
                    writing_perc = 10
                
                did_writing = True
            
            # Pdf creation:
            # initialise pdf and settings
            pdf = PDF()
            pdf.set_auto_page_break(auto = True)
            pdf.set_margin(15)
            pdf.add_font('cambria', '', 
                    'pdf_resources/Cambria.ttf',
                    uni = True)
            pdf.add_font('cambria', 'B', 
                    'pdf_resources/cambria-bold.ttf',
                    uni = True)
            pdf.add_font('cambria', 'I', 
                    'pdf_resources/cambria-italic.ttf',
                    uni = True)
            
            
            
            # Page 1: Summary page
            pdf.add_page()
            pdf.set_font('cambria', style = 'B', size = 23)
            pdf.cell(0, 10, test_name, 0, 1, 'C')
            pdf.set_font('cambria', '', size = 18)
            pdf.cell(0, 10, 'Student: {}'.format(name), 0, 1, 'C')
            pdf.ln(5)
            
            pdf.set_font('cambria', 'I', size = 14)
            pdf.multi_cell(0, 5, 
                        'The following report contains information on the {} ' \
                            'exams that the student has completed. As the student ' \
                                'did not complete every exam, the report is incomplete, including only ' \
                                    'data on the exams taken.'.format(np.sum((did_eng, did_maths, did_thinking, did_writing))))
            
            pdf.ln(3)
            
            pdf.set_font('cambria', '', size = 14)
            summary_table = [['Summary', 'Mark']]
            if did_eng:
                summary_table.append(['Reading', '{}/{}'.format(eng_mark, int(self.reading_total))])
            else:
                summary_table.append(['Reading', 'NA'])
            if did_maths:
                summary_table.append(['Mathematical Reasoning', '{}/{}'.format(maths_mark, int(self.maths_total))])
            else:
                summary_table.append(['Mathematical Reasoning', 'NA'])
            if did_thinking:
                summary_table.append(['Thinking Skills', '{}/{}'.format(thinking_mark, int(self.thinking_total))])
            else:
                summary_table.append(['Thinking Skills', 'NA'])
            if did_writing: 
                summary_table.append(['Writing', '{}/{}'.format(writing_mark, int(self.writing_total))])
            else:
                summary_table.append(['Writing', 'NA'])
            
            summary_table.append(['Overall Rank', 'NA'])
            
            pdf.create_table(table_data = summary_table, data_size = 14, cell_width = [150, 30],  
                            emphasize_data = ['Marks', 'Summary'], 
                            emphasize_style = 'B')
            
            pdf.ln(3)
            if did_eng:
                pdf.image('percentile_bands/reading' + str(eng_perc) + '.png', w = 200, x= 0.5)
            if did_maths:
                pdf.image('percentile_bands/maths' + str(maths_perc) + '.png', w = 200, x= 0.5)
            if did_thinking:
                pdf.image('percentile_bands/thinking' + str(thinking_perc) + '.png', w = 200, x= 0.5)
            if did_writing:
                pdf.image('percentile_bands/writing' + str(writing_perc) + '.png', w = 200, x= 0.5)
            
            
            
            # Page 2 & 3: Reading Comprehension
            if did_eng:
                pdf.add_page()
                pdf.new_section('Reading')
                pdf.ln(5)
            
                # table of statistics
                eng_table = [['Student Mark:', '{}/{}'.format(int(eng_mark), self.reading_total)],
                            ['Student Rank:', '{}/{}'.format(int(eng_rank), self.num_reading)],
                            ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.reading_mark), self.reading_total)],
                            ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.reading_mark)), self.reading_total)]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(eng_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                eng_fig = plots.reading_chart(self.agg_data, eng_mark)
                eng_fig.savefig('eng_fig.png', dpi = 500)
                plt.close(eng_fig)
                pdf.image('eng_fig.png', x = 30, w = 150)
                
                # table of question breakdown
                eng_q_table = [['Question Type', 'Questions', 'Mark'],
                            ['Normal Texts', 'Q1-14', '{}/14'.format(int(np.sum(student_info[12:26])))],
                            ['Cloze Passages', 'Q15-20', '{}/6'.format(int(np.sum(student_info[26:32])))],
                            ['Comparing Extracts', 'Q21-30', '{}/10'.format(int(np.sum(student_info[32:42])))]]
                pdf.create_table(eng_q_table, title = 'Question Breakdown', data_size = 14, title_size = 15, 
                                emphasize_data = ['Question Type', 'Questions', 'Mark'], emphasize_style = 'B',
                                cell_width = [100, 60, 20])
                
                # every reading comp question
                pdf.add_page()
                pdf.set_font('cambria', 'B', size = 14)
                pdf.cell(10, txt = 'Reading Test Breakdown')
                pdf.ln(5)
                pdf.set_font('cambria', '', 10)
                pdf.cell(0, 5, txt = 'Questions highlighted green indicate that the student responded correctly. Red highlight indicates an incorrect response.')
                pdf.ln(5)
                reading_qtable = pdf.questions_table('reading_q_types.xlsx', self.agg_data.iloc[:, 12:42])
                pdf.create_question_table(reading_qtable, self.agg_data.iloc[student_index, 12:42], did_test = True)
            
                
            else:
                pdf.add_page()
                pdf.new_section('Reading')
                pdf.ln(5)
            
                # table of statistics
                eng_table = [['Student Mark:', 'NA'],
                            ['Student Rank:', 'NA'],
                            ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.reading_mark), self.reading_total)],
                            ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.reading_mark)), self.reading_total)]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(eng_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                eng_fig = plots.reading_chart(self.agg_data, -1)
                eng_fig.savefig('eng_fig.png', dpi = 500)
                plt.close(eng_fig)
                pdf.image('eng_fig.png', x = 30, w = 150)
            
                # every reading comp question
                pdf.add_page()
                pdf.set_font('cambria', 'B', size = 14)
                pdf.cell(10, txt = 'Reading Test Breakdown')
                pdf.ln(5)
                reading_qtable = pdf.questions_table('reading_q_types.xlsx', self.agg_data.iloc[:, 12:42])
                pdf.create_question_table(reading_qtable, self.agg_data.iloc[student_index, 12:42], did_test = False)
            
            
            # Page 4 & 5: Mathematical Reasoning
            if did_maths:
                pdf.add_page()
                pdf.new_section('Mathematical Reasoning')
                pdf.ln(5)
                
                # table of statistics
                maths_table = [['Student Mark:', '{}/{}'.format(int(maths_mark), int(self.maths_total))],
                            ['Student Rank:', '{}/{}'.format(int(maths_rank), int(self.num_maths))],
                            ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.maths_mark), int(self.maths_total))],
                            ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.maths_mark)), int(self.maths_total))]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(maths_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                maths_fig = plots.maths_chart(self.agg_data, maths_mark)
                maths_fig.savefig('maths_fig.png', dpi = 500)
                plt.close(maths_fig)
                pdf.image('maths_fig.png', x = 30, w = 150)
                
                # message to say no maths question breakdown
                pdf.set_font('cambria', style = 'I', size = 14)
                pdf.multi_cell(0, 5, 
                            txt = '* The maths question breakdown will come at a later date while we revamp the question classification system.')
                
                
                
                # new page for every Maths Question
                pdf.add_page()
                pdf.set_font('cambria', 'B', size = 14)
                pdf.cell(10, txt = 'Mathematics Test Breakdown')
                pdf.ln(5)
                pdf.set_font('cambria', '', 10)
                pdf.cell(0, 5, txt = 'Questions highlighted green indicate that the student responded correctly. Red highlight indicates an incorrect response.')
                pdf.ln(5)
                maths_qtable = pdf.questions_table('maths_q_types.xlsx', self.agg_data.iloc[:, 42:77])
                pdf.create_question_table(maths_qtable, self.agg_data.iloc[student_index, 42:77], did_test = True)
                
            else:
                # Page 3 : Mathematics
                pdf.add_page()
                pdf.new_section('Mathematical Reasoning')
                pdf.ln(5)
                
                # table of statistics
                maths_table = [['Student Mark:', 'NA'],
                            ['Student Rank:', 'NA'],
                            ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.maths_mark), int(self.maths_total))],
                            ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.maths_mark)), int(self.maths_total))]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(maths_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                maths_fig = plots.maths_chart(self.agg_data, -1)
                maths_fig.savefig('maths_fig.png', dpi = 500)
                plt.close(maths_fig)
                pdf.image('maths_fig.png', x = 30, w = 150)
                
                
                # new page for every Maths Question
                pdf.add_page()
                pdf.set_font('cambria', 'B', size = 14)
                pdf.cell(10, txt = 'Mathematics Test Breakdown')
                pdf.ln(5)
                maths_qtable = pdf.questions_table('maths_q_types.xlsx', self.agg_data.iloc[:, 42:77])
                pdf.create_question_table(maths_qtable, self.agg_data.iloc[student_index, 42:77], did_test = False)
            
            # Page 6 & 7: Thinking Skills
            if did_thinking:
                pdf.add_page()
                pdf.new_section('Thinking Skills')
                pdf.ln(5)
                
                # table of statistics
                thinking_table = [['Student Mark:', '{}/{}'.format(int(thinking_mark), int(self.thinking_total))],
                                ['Student Rank:', '{}/{}'.format(int(thinking_rank), int(self.num_thinking))],
                                ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.thinking_mark), int(self.thinking_total))],
                                ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.thinking_mark)), int(self.thinking_total))]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(thinking_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                thinking_fig = plots.thinking_chart(self.agg_data, thinking_mark)
                thinking_fig.savefig('thinking_fig.png', dpi = 500)
                plt.close(thinking_fig)
                pdf.image('thinking_fig.png', x = 30, w = 150)
                
                
                # create table of common questions student got wrong
                thinking_q = pd.read_excel('thinking_q_types.xlsx').q_type
                top_tq = Counter(thinking_q).most_common()[0:4]
                t_marks = []
                for q in top_tq:
                    t_marks.append(np.sum(student_info.iloc[77:117].reset_index(drop = True)[thinking_q == q[0]]))
                
                t_common_q = [['Most Common Questions', 'Mark']]
                t_common_q2 = [[top_tq[i][0], '{}/{}'.format(int(t_marks[i]), top_tq[i][1])] for i in range(0, 4)]
                t_common_q = t_common_q + t_common_q2
                
                pdf.create_table(t_common_q, title = 'Question Breakdown', data_size = 14,
                                title_size = 15, emphasize_data = ['Most Common Questions', 'Mark'], emphasize_style = 'B',
                                cell_width = [150, 30])
                
                
                pdf.add_page()
                pdf.set_font('cambria', 'B', size = 14)
                pdf.cell(10, txt = 'Thinking Skills Test Breakdown')
                pdf.ln(5)
                pdf.set_font('cambria', '', 10)
                pdf.cell(0, 5, txt = 'Questions highlighted green indicate that the student responded correctly. Red highlight indicates an incorrect response.')
                pdf.ln(5)
                thinking_qtable = pdf.questions_table('thinking_q_types.xlsx', self.agg_data.iloc[:, 77:117])
                pdf.create_question_table(thinking_qtable, self.agg_data.iloc[student_index, 77:117], did_test = True)
            
            else:
                pdf.add_page()
                pdf.new_section('Thinking Skills')
                pdf.ln(5)
                
                # table of statistics
                thinking_table = [['Student Mark:', 'NA'],
                                ['Student Rank:', 'NA'],
                                ['Cohort Average:', '{:.2f}/{}'.format(np.nanmean(self.agg_data.thinking_mark), int(self.thinking_total))],
                                ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.thinking_mark)), int(self.thinking_total))]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(thinking_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                thinking_fig = plots.thinking_chart(self.agg_data, -1)
                thinking_fig.savefig('thinking_fig.png', dpi = 500)
                plt.close(thinking_fig)
                pdf.image('thinking_fig.png', x = 30, w = 150)
                
                # Page 6 : Every Thinking Skills Question
                pdf.add_page()
                pdf.set_font('cambria', 'B', size = 14)
                pdf.cell(10, txt = 'Thinking Skills Test Breakdown')
                pdf.ln(5)
                thinking_qtable = pdf.questions_table('thinking_q_types.xlsx', self.agg_data.iloc[:, 77:117])
                pdf.create_question_table(thinking_qtable, self.agg_data.iloc[student_index, 77:117], did_test = False)
            
            # create writing page if necessary
            if did_writing:
                pdf.add_page()
                pdf.new_section('Writing')
                pdf.ln(5)
                
                # table of statistics
                writing_table = [['Student Mark:', '{}/{}'.format(int(writing_mark), int(self.writing_total))],
                                ['Student Rank:', '{}/{}'.format(int(writing_rank), int(self.num_writing))],
                                ['Cohort Average:', '{:.2f}/{}'.format(np.mean(self.agg_data.writing_mark), self.writing_total)],
                                ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.writing_mark)), self.writing_total)]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(writing_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                writing_fig = plots.writing_chart(self.agg_data, writing_mark)
                writing_fig.savefig('writing_fig.png', dpi = 500)
                plt.close(writing_fig)
                pdf.image('writing_fig.png', x = 30, w = 150)
                
            else:
                pdf.add_page()
                pdf.new_section('Writing')
                pdf.ln(5)
                
                # table of statistics
                writing_table = [['Student Mark:', 'NA'],
                                ['Student Rank:', 'NA'],
                                ['Cohort Average:', '{:.2f}/{}'.format(np.mean(self.agg_data.writing_mark), self.writing_total)],
                                ['Maximum Mark:', '{}/{}'.format(int(np.max(self.agg_data.writing_mark)), self.writing_total)]]
                pdf.set_font('cambria', '', 14)
                pdf.create_mark_table(writing_table)
                
                pdf.ln(5)
                
                # create chart, and put in image
                writing_fig = plots.writing_chart(self.agg_data, -1)
                writing_fig.savefig('writing_fig.png', dpi = 500)
                plt.close(writing_fig)
                pdf.image('writing_fig.png', x = 30, w = 150)
                
            # output pdf
            pdf_name = '{}.pdf'.format(name)
            pdf.output(pdf_name, 'F')
            
            # create merged pdf with solutions
            merged_pdf = PdfMerger()
            merged_pdf.append(pdf_name)
            merged_pdf.append('sttc_sols.pdf')
            merged_pdf.write(pdf_name)
            merged_pdf.close()
            
            
            
        
    
if __name__ == '__main__':
    from tkinter.filedialog import askdirectory
    import os
    from data_preprocessing_SID import DataProcess
    
    path = askdirectory()
    os.chdir(path)
    
    # Load files
    sttc8_files = DataProcess('sttc', '1')
    sttc8_files.combine()
    reading_combined, maths_combined, thinking_combined, writing_combined = sttc8_files.diagnosis(output = True)
    
    # roll id
    rollid = pd.read_excel('sttc_rollid.xlsx') 
    
    # Report Creation Module
    report_creation = ReportCreation('sttc', '1', rollid, reading_combined, maths_combined, thinking_combined, writing_combined)
    report_creation.prepare()
    agg_data = report_creation.aggregate_data()
    
    