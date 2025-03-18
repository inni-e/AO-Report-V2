# Class for generating missing data and email excel sheets

import numpy as np
import pandas as pd

class MissingExcels():
    def __init__(self, test_type, test_no, rollid, reading_combined, maths_combined, thinking_combined, writing_combined = None):
        self.test_type = test_type
        self.test_no = test_no
        self.reading_combined = reading_combined
        self.maths_combined = maths_combined
        self.thinking_combined = thinking_combined
        self.rollid = rollid
        self.rollid.iloc[:, 1] = list(map(lambda x: x.strip(), self.rollid.iloc[:, 1]))
        
        if test_type != 'octt':
            self.writing_combined = writing_combined
        
    def missing_data(self):
        if self.test_type == 'octt':
            # get list of students who sat all exams
            complete_students = set(self.reading_combined.SID).intersection(self.maths_combined.SID).intersection(self.thinking_combined.SID)
            
            # create lists of incompletes that have each subject
            has_reading = self.reading_combined.loc[[not student in complete_students for student in self.reading_combined.SID], ].SID
            has_maths = self.maths_combined.loc[[not student in complete_students for student in self.maths_combined.SID], ].SID
            has_thinking = self.thinking_combined.loc[[not student in complete_students for student in self.thinking_combined.SID], ].SID
            
            # get list of students who have incomplete data
            incomplete_students = np.unique(np.concatenate([has_reading, has_maths, has_thinking]))
            
            # create excel dataframe
            self.missing_df = pd.DataFrame(columns = ['ID', 'Name', 'Branch', 'Reading', 'Maths', 'Thinking'])
            for i in range(len(incomplete_students)):
                student = incomplete_students[i]
                row = [student, '', '', '', '', '']
                no_info = True
                
                if not student in list(has_reading):
                    row[3] = 'X'
                else:
                    stu_ind = np.where(self.reading_combined.SID == student)[0]
                    row[1] = self.reading_combined.Name[stu_ind].values[0]
                    row[2] = self.reading_combined.centre[stu_ind].values[0]
                    no_info = False
                
                if not student in list(has_maths):
                    row[4] = 'X'
                elif no_info: 
                    stu_ind = np.where(self.maths_combined.SID == student)[0]
                    row[1] = self.maths_combined.Name[stu_ind].values[0]
                    row[2] = self.maths_combined.centre[stu_ind].values[0]
                    no_info = False
                    
                if not student in list(has_thinking):
                    row[5] = 'X'
                elif no_info: 
                    stu_ind = np.where(self.thinking_combined.SID == student)[0]
                    row[1] = self.thinking_combined.Name[stu_ind].values[0]
                    row[2] = self.thinking_combined.centre[stu_ind].values[0]
                
                # add row to dataframe
                self.missing_df.loc[i] = row
            
        else:
            # get list of students who sat all exams
            complete_students = set(self.reading_combined.SID).intersection(self.maths_combined.SID).intersection(self.thinking_combined.SID).intersection(self.writing_combined.SID)
            
            # create lists of incompletes that have each subject
            has_reading = self.reading_combined.loc[[not student in complete_students for student in self.reading_combined.SID], ].SID
            has_maths = self.maths_combined.loc[[not student in complete_students for student in self.maths_combined.SID], ].SID
            has_thinking = self.thinking_combined.loc[[not student in complete_students for student in self.thinking_combined.SID], ].SID
            has_writing = self.writing_combined.loc[[not student in complete_students for student in self.writing_combined.SID], ].SID
            
            # get list of students who have incomplete data
            incomplete_students = np.unique(np.concatenate([has_reading, has_maths, has_thinking, has_writing]))
            
            # create excel dataframe
            self.missing_df = pd.DataFrame(columns = ['SID', 'Name', 'Branch', 'Reading', 'Maths', 'Thinking', 'Writing'])
            for i in range(len(incomplete_students)):
                student = incomplete_students[i]
                row = [student, '', '', '', '', '', '']
                no_info = True
                if not student in list(has_reading):
                    row[3] = 'X'
                else:
                    stu_ind = np.where(self.reading_combined.SID == student)[0]
                    row[1] = self.reading_combined.Name[stu_ind].values[0]
                    row[2] = self.reading_combined.centre[stu_ind].values[0]
                    no_info = False
                
                if not student in list(has_maths):
                    row[4] = 'X'
                elif no_info: 
                    stu_ind = np.where(self.maths_combined.SID == student)[0]
                    row[1] = self.maths_combined.Name[stu_ind].values[0]
                    row[2] = self.maths_combined.centre[stu_ind].values[0]
                    no_info = False
                    
                if not student in list(has_thinking):
                    row[5] = 'X'
                elif no_info: 
                    stu_ind = np.where(self.thinking_combined.SID == student)[0]
                    row[1] = self.thinking_combined.Name[stu_ind].values[0]
                    row[2] = self.thinking_combined.centre[stu_ind].values[0]
                    no_info = False
                
                if not student in list(has_writing):
                    row[6] = 'X'
                elif no_info:
                    stu_ind = np.where(self.writing_combined.SID == student)[0]
                    row[1] = self.writing_combined.Name[stu_ind].values[0]
                    row[2] = self.writing_combined.centre[stu_ind].values[0]
                
                # add row
                self.missing_df.loc[i] = row
    
        return(self.missing_df)
        
    def missing_emails(self):
        # Students in Reading with missing emails
        reading_ind = [not id in list(self.rollid.iloc[:, 0]) for id in self.reading_combined.SID]
        missing_reading = self.reading_combined.loc[reading_ind, 'SID'].values
        
        # Students in Maths with missing emails
        maths_ind = [not id in list(self.rollid.iloc[:, 0]) for id in self.maths_combined.SID]
        missing_maths = self.maths_combined.loc[maths_ind, 'SID'].values
        
        # Students in Thinking Skills with missing emails
        thinking_ind = [not id in list(self.rollid.iloc[:, 0]) for id in self.thinking_combined.SID]
        missing_thinking = self.thinking_combined.loc[thinking_ind, 'SID'].values
        
        if self.test_type != 'octt':
            # Students in Writing with missing emails
            writing_ind = [not id in list(self.rollid.iloc[:, 0]) for id in self.writing_combined.SID]
            missing_writing = self.writing_combined.loc[writing_ind, 'SID'].values
            
            # Unique students with missing email
            missing_emails = np.unique(np.concatenate([missing_reading, missing_maths, missing_thinking, missing_writing]))
        else:
            missing_emails = np.unique(np.concatenate([missing_reading, missing_maths, missing_thinking]))
        
        # Get names for ids
        names = ['']*len(missing_emails)
        for i in range(len(missing_emails)):
            id = missing_emails[i]
            if id in list(self.reading_combined.SID):
                ind = np.where(self.reading_combined.SID == id)[0][0]
                names[i] = self.reading_combined.Name[ind]
                continue
            if id in list(self.maths_combined.SID):
                ind = np.where(self.maths_combined.SID == id)[0][0]
                names[i] = self.maths_combined.Name[ind]
                continue
            if id in list(self.thinking_combined.SID):
                ind = np.where(self.thinking_combined.SID == id)[0][0]
                names[i] = self.thinking_combined.Name[ind]
                continue
            if self.test_type != 'octt':
                ind = np.where(self.writing_combined.SID == id)[0][0]
                names[i] = self.writing_combined.Name[ind]
        
        # Create missing emails dataframe to output
        self.missing_email_df = pd.DataFrame({'Student ID' : missing_emails,
                                              'Name' : names,
                                              'Emails' : np.repeat('', len(missing_emails))})
        
        return(self.missing_email_df)
        



if __name__ == '__main__':
    from tkinter.filedialog import askdirectory
    import os
    from data_preprocessing_SID import DataProcess
    path = askdirectory()
    os.chdir(path)
    
    # Load files
    sttc8_files = DataProcess('sttc', '8')
    sttc8_files.combine()
    reading_combined, maths_combined, thinking_combined, writing_combined = sttc8_files.diagnosis(output = True)
    
    rollid = pd.read_excel('sttc_rollid.xlsx')
    missing_excels = MissingExcels('sttc', '8', rollid, reading_combined, maths_combined, thinking_combined, writing_combined)
    
    missing_data = missing_excels.missing_data()
    missing_emails = missing_excels.missing_emails()
    