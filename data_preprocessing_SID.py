# Class for importing data (FOR STUDENT ID SYSTEM)

import pandas as pd
import numpy as np

class DataProcess():
    def __init__(self, test_type, test_no, data_type):
        self.test_no = test_no
        self.og_type = test_type
        test_type = test_type[0:4]
        self.test_type = test_type
        self.data_type = data_type

        # load data
        # reading
        if self.data_type != 'online':
            try:
                self.bv_reading = pd.read_csv('bv_' + test_type + '_reading' + test_no + '.csv')
            except:
                try:
                    self.bv_reading = pd.read_excel('bv_' + test_type + '_reading' + test_no + '.xlsx')
                except:
                    raise NameError('Bella Vista reading file not found.')

            try:
                self.bur_reading = pd.read_csv('bur_' + test_type + '_reading' + test_no + '.csv')
            except:
                try: 
                    self.bur_reading = pd.read_excel('bur_' + test_type + '_reading' + test_no + '.xlsx')
                except:
                    raise NameError('Burwood reading file not found.')
            
            try: 
                self.epping_reading = pd.read_csv('epping_' + test_type + '_reading' + test_no + '.csv')
            except:
                try:
                    self.epping_reading = pd.read_excel('epping_' + test_type + '_reading' + test_no + '.xlsx')
                except:
                    raise NameError('Epping reading file not found.')
            
            try:
                self.parra_reading = pd.read_csv('parra_' + test_type + '_reading' + test_no + '.csv')
            except:
                try:
                    self.parra_reading = pd.read_excel('parra_' + test_type + '_reading' + test_no + '.xlsx')
                except:
                    raise NameError('Parramatta reading file not found.')
        
        if self.data_type != 'inperson':
            try:
                self.flexi_reading = pd.read_csv('flexi_' + test_type + '_reading' + test_no + '.csv')
            except:
                try:
                    self.flexi_reading = pd.read_excel('flexi_' + test_type + '_reading' + test_no + '.xlsx')
                except:
                    raise NameError('Flexiquiz reading file not found.')
        
        # maths
        if self.data_type != 'online':
            try:
                self.bv_maths = pd.read_csv('bv_' + test_type + '_maths' + test_no + '.csv')
            except: 
                try:
                    self.bv_maths = pd.read_excel('bv_' + test_type + '_maths' + test_no + '.xlsx')
                except:
                    raise NameError('Bella Vista maths file not found.')
                    
            try:
                self.bur_maths = pd.read_csv('bur_' + test_type + '_maths' + test_no + '.csv')
            except:
                try:
                    self.bur_maths = pd.read_excel('bur_' + test_type + '_maths' + test_no + '.xlsx')
                except:
                    raise NameError('Burwood maths file not found.')

            try:
                self.epping_maths = pd.read_csv('epping_' + test_type + '_maths' + test_no + '.csv')
            except:
                try:
                    self.epping_maths = pd.read_excel('epping_' + test_type + '_maths' + test_no + '.xlsx')
                except:
                    raise NameError('Epping maths file not found.')

            try:
                self.parra_maths = pd.read_csv('parra_' + test_type + '_maths' + test_no + '.csv')
            except:
                try:
                    self.parra_maths = pd.read_excel('parra_' + test_type + '_maths' + test_no + '.xlsx')
                except:
                    raise NameError('Parramatta maths file not found.')

        if self.data_type != 'inperson':
            try:
                self.flexi_maths = pd.read_csv('flexi_' + test_type + '_maths' + test_no + '.csv')
            except:
                try:
                    self.flexi_maths = pd.read_excel('flexi_' + test_type + '_maths' + test_no + '.xlsx')
                except:
                    raise NameError('Flexiquiz maths file not found.')
        
        # thinking
        if self.data_type != 'online':
            try:
                self.bv_thinking = pd.read_csv('bv_' + test_type + '_thinking' + test_no + '.csv')
            except:
                try:
                    self.bv_thinking = pd.read_excel('bv_' + test_type + '_thinking' + test_no + '.xlsx')
                except:
                    raise NameError('Bella Vista thinking skills file not found.')
            try:
                self.bur_thinking = pd.read_csv('bur_' + test_type + '_thinking' + test_no + '.csv')
            except:
                try:
                    self.bur_thinking = pd.read_excel('bur_' + test_type + '_thinking' + test_no + '.xlsx')
                except:
                    raise NameError('Burwood thinking skills file not found.')
            try:
                self.epping_thinking = pd.read_csv('epping_' + test_type + '_thinking' + test_no + '.csv')
            except:
                try:
                    self.epping_thinking = pd.read_excel('epping_' + test_type + '_thinking' + test_no + '.xlsx')
                except:
                    raise NameError('Epping thinking skills file not found.')

            try:
                self.parra_thinking = pd.read_csv('parra_' + test_type + '_thinking' + test_no + '.csv')
            except:
                try:
                    self.parra_thinking = pd.read_excel('parra_' + test_type + '_thinking' + test_no + '.xlsx')
                except:
                    raise NameError('Parramatta thinking skills file not found.')

        if self.data_type != 'inperson':
            try:
                self.flexi_thinking = pd.read_csv('flexi_' + test_type + '_thinking' + test_no + '.csv')
            except:
                try:
                    self.flexi_thinking = pd.read_excel('flexi_' + test_type + '_thinking' + test_no + '.xlsx')
                except:
                    raise NameError('Flexiquiz thinking skills file not found.')
            

        if test_type !='octt':
            # import writing files
            if self.data_type != 'online':
                try:
                    self.bv_writing = pd.read_csv('bv_' + test_type + '_writing' + test_no + '.csv')
                except:
                    try:
                        self.bv_writing = pd.read_excel('bv_' + test_type + '_writing' + test_no + '.xlsx')
                    except:
                        raise NameError('Bella Vista writing file not found.')

                try:
                    self.bur_writing = pd.read_csv('bur_' + test_type + '_writing' + test_no + '.csv')
                except:
                    try:
                        self.bur_writing = pd.read_excel('bur_' + test_type + '_writing' + test_no + '.xlsx')
                    except:
                        raise NameError('Burwood writing file not found.')
                
                try:
                    self.epping_writing = pd.read_csv('epping_' + test_type + '_writing' + test_no + '.csv')
                except:
                    try:
                        self.epping_writing = pd.read_excel('epping_' + test_type + '_writing' + test_no + '.xlsx')
                    except:
                        raise NameError('Epping writing file not found.')

                try:
                    self.parra_writing = pd.read_csv('parra_' + test_type + '_writing' + test_no + '.csv')
                except:
                    try:
                        self.parra_writing = pd.read_excel('parra_' + test_type + '_writing' + test_no + '.xlsx')
                    except:
                        raise NameError('Parramatta writing file not found.')

            if self.data_type != 'inperson':
                try:
                    self.flexi_writing = pd.read_csv('flexi_' + test_type + '_writing' + test_no + '.csv')
                except:
                    try:
                        self.flexi_writing = pd.read_excel('flexi_' + test_type + '_writing' + test_no + '.xlsx')
                    except:
                        raise NameError('Flexiquiz writing file not found.')
        
        
    ########################################
    ##--------- HELPER FUNCTIONS ---------##
    ########################################
    def flexi_extract(self, dataset, test_type, is_eng = False, ans_sheet = None):
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
    
    def clean_duplicates(self, dataset, subject):
        ## ----------- Catch duplicates in phone sheet ---------------- ##
        normal_students = np.where(dataset['centre'] == 'online')[0][0]
        uniques = np.unique(dataset.iloc[0:normal_students, 2], return_counts = True)
        if any(uniques[1] > 1):
            duplicated_ids = list(uniques[0][uniques[1] > 1])
            raise EnvironmentError('Duplicated student IDs in ' + subject + ': ' + str(duplicated_ids))
        
        ## ---------------- Catch flexi duplicates ------------------ ##
        uniques = np.unique(dataset.iloc[normal_students:, 2], return_counts = True)
        if any(uniques[1] > 1):
            duplicated_ids = list(uniques[0][uniques[1] > 1])
            raise EnvironmentError('Duplicates student IDs in flexiquiz ' + subject + ': ' + str(duplicated_ids))
        
        ## --------------- Remoe duplicates across formats ------------ ##
        repeated_ind = [id in list(dataset.iloc[0:normal_students, 2]) for id in dataset.iloc[normal_students:, 2]]
        wanted_vals = np.append(np.repeat(True, normal_students), ~np.array(repeated_ind))
        dataset = dataset.iloc[wanted_vals, :].reset_index(drop = True)
        
        return(dataset)
        
        
        
    ###################################################
    ### ------------ Data Processing --------------- ##
    ###################################################      
    def combine(self, output = False): 
        if self.data_type == 'hybrid':
            # FILE PREPARATION ------------------
            # Prepare Reading Files
            # raise errors
            if not (self.bv_reading.shape[1] == self.bur_reading.shape[1] == self.epping_reading.shape[1] == self.parra_reading.shape[1]):
                raise ValueError('Number of columns in reading files do not match.')
            
            if not self.bv_reading.iloc[:, 2].is_unique:
                raise ValueError('Repeated Student Ids in BV reading file.')

            if not self.bur_reading.iloc[:, 2].is_unique:
                raise ValueError('Repeated Student Ids in Burwood reading file.')
            
            if not self.epping_reading.iloc[:, 2].is_unique:
                raise ValueError('Repeated student ids in Epping reading file.')
            
            if not self.parra_reading.iloc[:, 2].is_unique:
                raise ValueError('Repeated student ids in Parramatta reading file.')
            
            # label students by branch
            self.bv_reading.iloc[:, 0] = np.repeat('bv', self.bv_reading.shape[0])
            self.bur_reading.iloc[:, 0] = np.repeat('bur', self.bur_reading.shape[0])
            self.epping_reading.iloc[:, 0] = np.repeat('epping', self.epping_reading.shape[0])
            self.parra_reading.iloc[:, 0] = np.repeat('parra', self.parra_reading.shape[0])
            
            # ensure column names are all equal
            reading_colnames = list(self.bv_reading.columns)
            reading_colnames[0] = 'centre'
            self.bv_reading.columns = reading_colnames
            self.bur_reading.columns = reading_colnames
            self.epping_reading.columns = reading_colnames
            self.parra_reading.columns = reading_colnames

            # create a placeholder flexi dataframe to append to data
            flexi_placeholder_reading = pd.DataFrame(np.tile(np.nan, (self.flexi_reading.shape[0]-1, self.bv_reading.shape[1])))
            flexi_placeholder_reading.columns = reading_colnames

            # Prepare Maths Files
            # label students by branch
            self.bv_maths.iloc[:, 0] = np.repeat('bv', self.bv_maths.shape[0])
            self.bur_maths.iloc[:, 0] = np.repeat('bur', self.bur_maths.shape[0])
            self.epping_maths.iloc[:, 0] = np.repeat('epping', self.epping_maths.shape[0])
            self.parra_maths.iloc[:, 0] = np.repeat('parra', self.parra_maths.shape[0])

            # raise errors
            if not (self.bv_maths.shape[1] == self.bur_maths.shape[1] == self.epping_maths.shape[1] == self.parra_maths.shape[1]):
                raise ValueError('Number of columns in maths files do not match.')
            
            if not self.bv_maths.iloc[:, 2].is_unique:
                raise ValueError('Repeated Student Ids in BV maths file.')

            if not self.bur_maths.iloc[:, 2].is_unique:
                raise ValueError('Repeated Student Ids in Burwood maths file.')
            
            if not self.epping_maths.iloc[:, 2].is_unique:
                raise ValueError('Repeated student ids in Epping maths file.')
            
            if not self.parra_maths.iloc[:, 2].is_unique:
                raise ValueError('Repeated student ids in Parramatta maths file.')
            
            # ensure columns have the same names
            maths_colnames = list(self.bv_maths.columns)
            maths_colnames[0] = 'centre'
            self.bv_maths.columns = maths_colnames
            self.bur_maths.columns = maths_colnames
            self.epping_maths.columns = maths_colnames
            self.parra_maths.columns = maths_colnames

            # create placeholder flexi dataframe
            flexi_placeholder_maths = pd.DataFrame(np.tile(np.nan, (self.flexi_maths.shape[0], self.bv_maths.shape[1])))
            flexi_placeholder_maths.columns = maths_colnames

            # Prepare Thinking Skill Files
            # label students by branch
            self.bv_thinking.iloc[:, 0] = np.repeat('bv', self.bv_thinking.shape[0])
            self.bur_thinking.iloc[:, 0] = np.repeat('bur', self.bur_thinking.shape[0])
            self.epping_thinking.iloc[:, 0] = np.repeat('epping', self.epping_thinking.shape[0])
            self.parra_thinking.iloc[:, 0] = np.repeat('parra', self.parra_thinking.shape[0])

            if not (self.bv_thinking.shape[1] == self.bur_thinking.shape[1] == self.epping_thinking.shape[1] == self.parra_thinking.shape[1]):
                raise ValueError('Number of columns in thinking files do not match.')
            
            if not self.bv_thinking.iloc[:, 2].is_unique:
                raise ValueError('Repeated Student Ids in BV thinking file.')

            if not self.bur_thinking.iloc[:, 2].is_unique:
                raise ValueError('Repeated Student Ids in Burwood thinking file.')
            
            if not self.epping_thinking.iloc[:, 2].is_unique:
                raise ValueError('Repeated student ids in Epping thinking file.')
            
            if not self.parra_thinking.iloc[:, 2].is_unique:
                raise ValueError('Repeated student ids in Parramatta thinking file.')
            
            # ensure columns have the same names
            thinking_colnames = list(self.bv_thinking.columns)
            thinking_colnames[0] = 'centre'
            self.bv_thinking.columns = thinking_colnames
            self.bur_thinking.columns = thinking_colnames
            self.epping_thinking.columns = thinking_colnames
            self.parra_thinking.columns = thinking_colnames

            # create placeholder flexi dataframe
            flexi_placeholder_thinking = pd.DataFrame(np.tile(np.nan, (self.flexi_thinking.shape[0], self.bv_thinking.shape[1])))
            flexi_placeholder_thinking.columns = thinking_colnames

            # Prepare Writing Files
            # ensure each file only has 3 columns
            if self.test_type != 'octt':
                if self.bv_writing.shape[1] != 3:
                    self.bv_writing = self.bv_writing.iloc[:, 0:3]
                if self.bur_writing.shape[1] != 3:
                    self.bur_writing = self.bur_writing.iloc[:, 0:3]
                if self.epping_writing.shape[1] != 3:
                    self.epping_writing = self.epping_writing.iloc[:, 0:3]
                if self.parra_writing.shape[1] != 3:
                    self.parra_writing = self.parra_writing.iloc[:, 0:3]
                    
                # throw up any errors
                if not self.bv_writing.iloc[:, 1].is_unique:
                    raise ValueError('Repeated Student Ids in BV writing file.')

                if not self.bur_writing.iloc[:, 1].is_unique:
                    raise ValueError('Repeated Student Ids in Burwood writing file.')
                
                if not self.epping_writing.iloc[:, 1].is_unique:
                    raise ValueError('Repeated student ids in Epping writing file.')
                
                if not self.parra_writing.iloc[:, 1].is_unique:
                    raise ValueError('Repeated student ids in Parramatta writing file.')
                
                # add centre column
                self.bv_writing['centre'] = np.repeat('bv', self.bv_writing.shape[0])
                self.bur_writing['centre'] = np.repeat('bur', self.bur_writing.shape[0])
                self.epping_writing['centre'] = np.repeat('epping', self.epping_writing.shape[0])
                self.parra_writing['centre'] = np.repeat('parra', self.parra_writing.shape[0])

                # rename columns
                writing_colnames = ['Name', 'SID', 'Points', 'centre']
                self.bv_writing.columns = writing_colnames
                self.bur_writing.columns = writing_colnames
                self.epping_writing.columns = writing_colnames
                self.parra_writing.columns = writing_colnames

                # create flexi dataframe for writing
                flexi_placeholder_writing = pd.DataFrame(np.tile(np.nan, (self.flexi_writing.shape[0], self.bv_writing.shape[1])))
                flexi_placeholder_writing.columns = writing_colnames

            # FILE COMBINATION -------------------
            # Combine Reading Files (NOTE: First row of flexi reading is not a student entry)
            self.reading_combined = pd.concat([self.bv_reading, self.bur_reading, self.epping_reading, self.parra_reading,
                                            flexi_placeholder_reading], ignore_index = True)
            
            # Add flexi component
            # get flexi names
            flexi_reading_names = list(self.flexi_reading.iloc[1:, 0])
            self.reading_combined.iloc[-len(flexi_reading_names):, 3] = flexi_reading_names
            
            # add flexi ids
            flexi_reading_ids = list(self.flexi_reading.iloc[1:, 1])
            self.reading_combined.iloc[-len(flexi_reading_ids):, 2] = flexi_reading_ids
            
            # get flexi marks
            self.reading_combined.iloc[-len(flexi_reading_names):, 4] = self.flexi_reading.iloc[1:, 4]
            try:
                reading_ans_sheet = pd.read_excel('reading_q_types.xlsx')
            except:
                reading_ans_sheet = pd.read_csv('reading_q_types.csv')
            reading_ans = self.flexi_extract(dataset = self.flexi_reading.iloc[1:, ], test_type = self.test_type, is_eng = True, ans_sheet = reading_ans_sheet)
            for i in range(0, reading_ans.shape[1]):
                self.reading_combined.iloc[-len(flexi_reading_names):, 13+3*i] = pd.Series(reading_ans[:, i], dtype=self.reading_combined.iloc[:, 13+3*i].dtype)
            
            # add centre tag
            self.reading_combined.iloc[-len(flexi_reading_names):, 0] = 'online'
            
            # Fix columns
            self.reading_combined['Name'] = list(map(lambda x: x.strip(), self.reading_combined['Name']))
            
            
            # Combine Maths Files
            self.maths_combined = pd.concat([self.bv_maths, self.bur_maths, self.epping_maths, self.parra_maths, 
                                            flexi_placeholder_maths], ignore_index = True)
            
            # Add flexi component
            # get flexi names
            flexi_maths_names = list(self.flexi_maths.iloc[:, 0])
            self.maths_combined.iloc[-len(flexi_maths_names):, 3] = flexi_maths_names
            
            # get flexi ids
            flexi_maths_ids = list(self.flexi_maths.iloc[:, 1])
            self.maths_combined.iloc[-len(flexi_maths_ids):, 2] = flexi_maths_ids
            
            # get flexi marks
            self.maths_combined.iloc[-len(flexi_maths_names):, 4] = self.flexi_maths.iloc[:, 4]
            maths_ans = self.flexi_extract(self.flexi_maths, self.test_type)
            for i in range(0, maths_ans.shape[1]):
                self.maths_combined.iloc[-len(flexi_maths_names):, 13+3*i] = pd.Series(maths_ans[:, i], dtype=self.maths_combined.iloc[:, 13+3*i].dtype)
            
            # add centre tag
            self.maths_combined.iloc[-len(flexi_maths_names):, 0] = 'online'
            
            # Fix columns
            self.maths_combined['Name'] = list(map(lambda x: x.strip(), self.maths_combined['Name']))
            
            
            # Combine Thinking Skills Files
            self.thinking_combined = pd.concat([self.bv_thinking, self.bur_thinking, self.epping_thinking, self.parra_thinking,
                                                flexi_placeholder_thinking], ignore_index = True)
            # add flexi component
            # flexi names
            flexi_thinking_names = list(self.flexi_thinking.iloc[:, 0])
            self.thinking_combined.iloc[-len(flexi_thinking_names):, 3] = flexi_thinking_names
            
            # flexi ids
            flexi_thinking_ids = list(self.flexi_thinking.iloc[:, 1])
            self.thinking_combined.iloc[-len(flexi_thinking_ids):, 2] = flexi_thinking_ids
            
            # flexi marks
            self.thinking_combined.iloc[-len(flexi_thinking_names):, 4] = self.flexi_thinking.iloc[:, 4]
            thinking_ans = self.flexi_extract(self.flexi_thinking, self.test_type)
            for i in range(0, thinking_ans.shape[1]):
                self.thinking_combined.iloc[-len(flexi_thinking_names):, 13+3*i] = pd.Series(thinking_ans[:, i], dtype=self.thinking_combined.iloc[:, 13+3*i].dtype)
            
            # add centre tag
            self.thinking_combined.iloc[-len(flexi_thinking_names):, 0] = 'online'
            
            # Fix columns
            self.thinking_combined['Name'] = list(map(lambda x: x.strip(), self.thinking_combined['Name']))
        
        
            # Combine Writing Files
            if self.test_type != 'octt':
                self.writing_combined = pd.concat([self.bv_writing, self.bur_writing, self.epping_writing, self.parra_writing, 
                                                flexi_placeholder_writing], ignore_index = True)
                
                # add flexi writing names
                flexi_writing_names = self.flexi_writing.iloc[:, 0]
                self.writing_combined.iloc[-len(flexi_writing_names):, 0] = flexi_writing_names
                
                # add flexi ids
                flexi_writing_ids = self.flexi_writing.iloc[:, 1]
                self.writing_combined.iloc[-len(flexi_writing_ids):, 1] = flexi_writing_ids
                
                # add flexi writing marks
                self.writing_combined.iloc[-len(flexi_writing_names):, 2] = self.flexi_writing.iloc[:, 4]
                
                # add centre tag
                self.writing_combined.iloc[-len(flexi_writing_names):, 3] = 'online'
                
                # Fix columns
                self.writing_combined['Name'] = list(map(lambda x: x.strip(), self.writing_combined['Name']))
            
            # return combined dataframes
            if output:
                if self.test_type == 'octt':
                    return(self.reading_combined, self.maths_combined, self.thinking_combined)
                else:
                    return(self.reading_combined, self.maths_combined, self.thinking_combined, self.writing_combined)
        
        
        
        
        elif self.data_type == 'inperson':
            # FILE PREPARATION ------------------
            # Prepare Reading Files
            # raise errors
            if not (self.bv_reading.shape[1] == self.bur_reading.shape[1] == self.epping_reading.shape[1] == self.parra_reading.shape[1]):
                raise ValueError('Number of columns in reading files do not match.')
            
            if not self.bv_reading.iloc[:, 2].is_unique:
                raise ValueError('Repeated Student Ids in BV reading file.')

            if not self.bur_reading.iloc[:, 2].is_unique:
                raise ValueError('Repeated Student Ids in Burwood reading file.')
            
            if not self.epping_reading.iloc[:, 2].is_unique:
                raise ValueError('Repeated student ids in Epping reading file.')
            
            if not self.parra_reading.iloc[:, 2].is_unique:
                raise ValueError('Repeated student ids in Parramatta reading file.')
            
            # label students by branch
            self.bv_reading.iloc[:, 0] = np.repeat('bv', self.bv_reading.shape[0])
            self.bur_reading.iloc[:, 0] = np.repeat('bur', self.bur_reading.shape[0])
            self.epping_reading.iloc[:, 0] = np.repeat('epping', self.epping_reading.shape[0])
            self.parra_reading.iloc[:, 0] = np.repeat('parra', self.parra_reading.shape[0])
            
            # ensure column names are all equal
            reading_colnames = list(self.bv_reading.columns)
            reading_colnames[0] = 'centre'
            self.bv_reading.columns = reading_colnames
            self.bur_reading.columns = reading_colnames
            self.epping_reading.columns = reading_colnames
            self.parra_reading.columns = reading_colnames


            # Prepare Maths Files
            # label students by branch
            self.bv_maths.iloc[:, 0] = np.repeat('bv', self.bv_maths.shape[0])
            self.bur_maths.iloc[:, 0] = np.repeat('bur', self.bur_maths.shape[0])
            self.epping_maths.iloc[:, 0] = np.repeat('epping', self.epping_maths.shape[0])
            self.parra_maths.iloc[:, 0] = np.repeat('parra', self.parra_maths.shape[0])

            # raise errors
            if not (self.bv_maths.shape[1] == self.bur_maths.shape[1] == self.epping_maths.shape[1] == self.parra_maths.shape[1]):
                raise ValueError('Number of columns in maths files do not match.')
            
            if not self.bv_maths.iloc[:, 2].is_unique:
                raise ValueError('Repeated Student Ids in BV maths file.')

            if not self.bur_maths.iloc[:, 2].is_unique:
                raise ValueError('Repeated Student Ids in Burwood maths file.')
            
            if not self.epping_maths.iloc[:, 2].is_unique:
                raise ValueError('Repeated student ids in Epping maths file.')
            
            if not self.parra_maths.iloc[:, 2].is_unique:
                raise ValueError('Repeated student ids in Parramatta maths file.')
            
            # ensure columns have the same names
            maths_colnames = list(self.bv_maths.columns)
            maths_colnames[0] = 'centre'
            self.bv_maths.columns = maths_colnames
            self.bur_maths.columns = maths_colnames
            self.epping_maths.columns = maths_colnames
            self.parra_maths.columns = maths_colnames


            # Prepare Thinking Skill Files
            # label students by branch
            self.bv_thinking.iloc[:, 0] = np.repeat('bv', self.bv_thinking.shape[0])
            self.bur_thinking.iloc[:, 0] = np.repeat('bur', self.bur_thinking.shape[0])
            self.epping_thinking.iloc[:, 0] = np.repeat('epping', self.epping_thinking.shape[0])
            self.parra_thinking.iloc[:, 0] = np.repeat('parra', self.parra_thinking.shape[0])

            if not (self.bv_thinking.shape[1] == self.bur_thinking.shape[1] == self.epping_thinking.shape[1] == self.parra_thinking.shape[1]):
                raise ValueError('Number of columns in thinking files do not match.')
            
            if not self.bv_thinking.iloc[:, 2].is_unique:
                raise ValueError('Repeated Student Ids in BV thinking file.')

            if not self.bur_thinking.iloc[:, 2].is_unique:
                raise ValueError('Repeated Student Ids in Burwood thinking file.')
            
            if not self.epping_thinking.iloc[:, 2].is_unique:
                raise ValueError('Repeated student ids in Epping thinking file.')
            
            if not self.parra_thinking.iloc[:, 2].is_unique:
                raise ValueError('Repeated student ids in Parramatta thinking file.')
            
            # ensure columns have the same names
            thinking_colnames = list(self.bv_thinking.columns)
            thinking_colnames[0] = 'centre'
            self.bv_thinking.columns = thinking_colnames
            self.bur_thinking.columns = thinking_colnames
            self.epping_thinking.columns = thinking_colnames
            self.parra_thinking.columns = thinking_colnames

            # Prepare Writing Files
            # ensure each file only has 3 columns
            if self.test_type != 'octt':
                if self.bv_writing.shape[1] != 3:
                    self.bv_writing = self.bv_writing.iloc[:, 0:3]
                if self.bur_writing.shape[1] != 3:
                    self.bur_writing = self.bur_writing.iloc[:, 0:3]
                if self.epping_writing.shape[1] != 3:
                    self.epping_writing = self.epping_writing.iloc[:, 0:3]
                if self.parra_writing.shape[1] != 3:
                    self.parra_writing = self.parra_writing.iloc[:, 0:3]
                    
                # throw up any errors
                if not self.bv_writing.iloc[:, 1].is_unique:
                    raise ValueError('Repeated Student Ids in BV writing file.')

                if not self.bur_writing.iloc[:, 1].is_unique:
                    raise ValueError('Repeated Student Ids in Burwood writing file.')
                
                if not self.epping_writing.iloc[:, 1].is_unique:
                    raise ValueError('Repeated student ids in Epping writing file.')
                
                if not self.parra_writing.iloc[:, 1].is_unique:
                    raise ValueError('Repeated student ids in Parramatta writing file.')
                
                # add centre column
                self.bv_writing['centre'] = np.repeat('bv', self.bv_writing.shape[0])
                self.bur_writing['centre'] = np.repeat('bur', self.bur_writing.shape[0])
                self.epping_writing['centre'] = np.repeat('epping', self.epping_writing.shape[0])
                self.parra_writing['centre'] = np.repeat('parra', self.parra_writing.shape[0])

                # rename columns
                writing_colnames = ['Name', 'SID', 'Points', 'centre']
                self.bv_writing.columns = writing_colnames
                self.bur_writing.columns = writing_colnames
                self.epping_writing.columns = writing_colnames
                self.parra_writing.columns = writing_colnames
        
            # File Combination
            self.reading_combined = pd.concat([self.bv_reading, self.bur_reading, self.epping_reading, self.parra_reading], ignore_index = True)
            self.maths_combined = pd.concat([self.bv_maths, self.bur_maths, self.epping_maths, self.parra_maths], ignore_index = True)
            self.thinking_combined = pd.concat([self.bv_thinking, self.bur_thinking, self.epping_thinking, self.parra_thinking], ignore_index = True)
            
            if self.test_type != 'octt':
                self.writing_combined = pd.concat([self.bv_writing, self.bur_writing, self.epping_writing, self.parra_writing], ignore_index = True)
            
            if output:
                if self.test_type == 'octt':
                    return(self.reading_combined, self.maths_combined, self.thinking_combined)
                else:
                    return(self.reading_combined, self.maths_combined, self.thinking_combined, self.writing_combined)
        
        
        elif self.data_type == 'online':
            self.reading_combined = self.flexi_reading.iloc[1:, :]
            self.maths_combined = self.flexi_maths
            self.thinking_combined = self.flexi_thinking
            if self.test_type != 'octt':
                self.writing_combined = self.flexi_writing
            
            if output:
                if self.test_type == 'octt':
                    return(self.reading_combined, self.maths_combined, self.thinking_combined)
                else:
                    return(self.reading_combined, self.maths_combined, self.thinking_combined, self.writing_combined)
            
            
        
    def diagnosis(self, output = False):
        if self.data_type == 'hybrid':
            ## --------------- Remove missing student ids ------------------- ##
            self.reading_combined = self.reading_combined.loc[~self.reading_combined.iloc[:, 2].isna(), ].reset_index(drop = True)
            self.maths_combined = self.maths_combined.loc[~self.maths_combined.iloc[:, 2].isna(), ].reset_index(drop = True)
            self.thinking_combined = self.thinking_combined.loc[~self.thinking_combined.iloc[:, 2].isna(), ].reset_index(drop = True)
            if self.test_type != 'octt':
                self.writing_combined = self.writing_combined.loc[~self.writing_combined.iloc[:, 1].isna(), ].reset_index(drop = True)
            
            # change ids to integer
            self.reading_combined.iloc[:, 2] = list(map(lambda x: int(x), self.reading_combined.iloc[:, 2]))
            self.maths_combined.iloc[:, 2] = list(map(lambda x: int(x), self.maths_combined.iloc[:, 2]))
            self.thinking_combined.iloc[:, 2] = list(map(lambda x: int(x), self.thinking_combined.iloc[:, 2]))
            if self.test_type != 'octt':
                self.writing_combined.iloc[:, 1] = list(map(lambda x: int(x), self.writing_combined.iloc[:, 1]))    
            
            ## -------------------- Catch duplicates ----------------------- ##
            self.reading_combined = self.clean_duplicates(self.reading_combined, 'Reading Comprehension')
            self.maths_combined = self.clean_duplicates(self.maths_combined, 'Mathematical Reasoning')
            self.thinking_combined = self.clean_duplicates(self.thinking_combined, 'Thinking Skills')
            
            # Rename roll no column 
            self.reading_combined = self.reading_combined.rename(columns = {'Roll No':'SID'})
            self.maths_combined = self.maths_combined.rename(columns = {'Roll No':'SID'})
            self.thinking_combined = self.thinking_combined.rename(columns = {'Roll No': 'SID'})
            
            ## --------------------- Clean writing file -------------------- ##
            if self.test_type != 'octt':
                remove_ind = np.array([])
                for i in range(0, self.writing_combined.shape[0]):
                    if pd.isna(self.writing_combined.iloc[i, 2]) or isinstance(self.writing_combined.iloc[i, 2], str):
                        remove_ind = np.append(remove_ind, i)
                self.writing_combined = self.writing_combined.drop(remove_ind)

                ## ----------- Catch marks outside range ---------------- ##
                if (self.og_type != 'wemtoc'):
                    out_of_range_indices = self.writing_combined[(self.writing_combined.iloc[:, 2] > 25) | (self.writing_combined.iloc[:, 2] < 0)].index
                else:
                    out_of_range_indices = self.writing_combined[(self.writing_combined.iloc[:, 2] > 20) | (self.writing_combined.iloc[:, 2] < 0)].index

                if not out_of_range_indices.empty:
                    # Collect tuples of (child's name, centre) for each out-of-range mark
                    out_of_range_info = [(self.writing_combined.iloc[index, 0], self.writing_combined.iloc[index, 3]) for index in out_of_range_indices]
                    raise EnvironmentError('Marks out of acceptable range detected in Writing: ' + str(out_of_range_info))
                
                ## ----------- Catch duplicates in phone sheet ---------------- ##
                normal_students = np.where(self.writing_combined['centre'] == 'online')[0][0]
                uniques = np.unique(self.writing_combined.iloc[0:normal_students, 1], return_counts = True)
                if any(uniques[1] > 1):
                    duplicated_ids = list(uniques[0][uniques[1] > 1])
                    raise EnvironmentError('Duplicated student IDs in Writing: ' + str(duplicated_ids))
                
                ## ---------------- Catch flexi duplicates ------------------ ##
                uniques = np.unique(self.writing_combined.iloc[normal_students:, 1], return_counts = True)
                if any(uniques[1] > 1):
                    duplicated_ids = list(uniques[0][uniques[1] > 1])
                    raise EnvironmentError('Duplicates student IDs in flexiquiz Writing: ' + str(duplicated_ids))
                
                ## --------------- Remoe duplicates across formats ------------ ##
                repeated_ind = [id in list(self.writing_combined.iloc[0:normal_students, 1]) for id in self.writing_combined.iloc[normal_students:, 1]]
                wanted_vals = np.append(np.repeat(True, normal_students), ~np.array(repeated_ind))
                self.writing_combined = self.writing_combined.iloc[wanted_vals, :].reset_index(drop = True)
            
            if output:
                if self.test_type == 'octt':
                    return(self.reading_combined, self.maths_combined, self.thinking_combined)
                if self.test_type != 'octt':
                    return(self.reading_combined, self.maths_combined, self.thinking_combined, self.writing_combined)
        
        elif self.data_type == 'inperson':
            ## --------------- Remove missing student ids ------------------- ##
            self.reading_combined = self.reading_combined.loc[~self.reading_combined.iloc[:, 2].isna(), ]
            self.maths_combined = self.maths_combined.loc[~self.maths_combined.iloc[:, 2].isna(), ]
            self.thinking_combined = self.thinking_combined.loc[~self.thinking_combined.iloc[:, 2].isna(), ]
            if self.test_type != 'octt':
                self.writing_combined = self.writing_combined.loc[~self.writing_combined.iloc[:, 2].isna(), ]
            
            # change ids to integer
            self.reading_combined.iloc[:, 2] = list(map(lambda x: int(x), self.reading_combined.iloc[:, 2]))
            self.maths_combined.iloc[:, 2] = list(map(lambda x: int(x), self.maths_combined.iloc[:, 2]))
            self.thinking_combined.iloc[:, 2] = list(map(lambda x: int(x), self.thinking_combined.iloc[:, 2]))
            if self.test_type != 'octt':
                self.writing_combined.iloc[:, 1] = list(map(lambda x: int(x), self.writing_combined.iloc[:, 1]))    
                
            ## -------------------------------- Catch Duplicates ------------------------------##
            # Reading Comprehension
            uniques = np.unique(self.reading_combined.iloc[:, 2], return_counts = True)
            if any(uniques[1] > 1):
                duplicated_ids = list(uniques[0][uniques[1]>1])
                raise EnvironmentError('Duplicated student IDs in Reading Comprehension: ' + str(duplicated_ids))
            
            # Mathematical Reasoning
            uniques = np.unique(self.maths_combined.iloc[:, 2], return_counts = True)
            if any(uniques[1] > 1):
                duplicated_ids = list(uniques[0][uniques[1]>1])
                raise EnvironmentError('Duplicated student IDs in Mathematical Reasoning: ' + str(duplicated_ids))
            
            # Thinking Combined
            uniques = np.unique(self.thinking_combined.iloc[:, 2], return_counts = True)
            if any(uniques[1] > 1):
                duplicated_ids = list(uniques[0][uniques[1]>1])
                raise EnvironmentError('Duplicated student IDs in Thinking Skills: ' + str(duplicated_ids))
            
            # Writing
            if self.test_type != 'octt':
                uniques = np.unique(self.writing_combined.iloc[:, 1], return_counts = True)
                if any(uniques[1] > 1):
                    duplicated_ids = list(uniques[0][uniques[1]>1])
                    raise EnvironmentError('Duplicated student IDs in Writing: ' + str(duplicated_ids))

                if (self.og_type != 'wemtoc'):
                    out_of_range_indices = self.writing_combined[(self.writing_combined.iloc[:, 2] > 25) | (self.writing_combined.iloc[:, 2] < 0)].index
                else:
                    out_of_range_indices = self.writing_combined[(self.writing_combined.iloc[:, 2] > 20) | (self.writing_combined.iloc[:, 2] < 0)].index

                if not out_of_range_indices.empty:
                    # Collect tuples of (child's name, centre) for each out-of-range mark
                    out_of_range_info = [(self.writing_combined.iloc[index, 0], self.writing_combined.iloc[index, 3]) for index in out_of_range_indices]
                    raise EnvironmentError('Marks out of acceptable range detected in Writing: ' + str(out_of_range_info))

            if output:
                if self.test_type == 'octt':
                    return(self.reading_combined, self.maths_combined, self.thinking_combined)
                else:
                    return(self.reading_combined, self.maths_combined, self.thinking_combined, self.writing_combined)

            

        

if __name__ == '__main__':
    from tkinter.filedialog import askdirectory
    import os
    path = askdirectory()
    os.chdir(path)
    
    # Load files
    sttc8_files = DataProcess('sttc', '8')
    sttc8_files.combine()
    reading_combined, maths_combined, thinking_combined, writing_combined = sttc8_files.diagnosis(output = True)
