# Module for Creating Pdf's 

from fpdf import FPDF
import pandas as pd
import numpy as np

class PDF(FPDF):
    def header(self):
        self.image('pdf_resources/logo.png', 10, 8, 65)
        self.ln(20)
    def footer(self):
        self.set_y(-15)
        self.set_font('helvetica', '', 10)
        self.cell(0, 10, f'{self.page_no()}', align = 'C')
        
    def new_section(self, course_name):
        self.set_font('cambria', 'B', 20)
        self.cell(0, 10, course_name, 0, 1, 'C')
        
    
    def create_table(self, table_data, title='', data_size = 10, title_size=12, align_data='L', align_header='L', cell_width='even', x_start='x_default',emphasize_data=[], emphasize_style=None,emphasize_color=(0,0,0)): 
        """
        table_data: 
                    list of lists with first element being list of headers
        title: 
                    (Optional) title of table (optional)
        data_size: 
                    the font size of table data
        title_size: 
                    the font size fo the title of the table
        align_data: 
                    align table data
                    L = left align
                    C = center align
                    R = right align
        align_header: 
                    align table data
                    L = left align
                    C = center align
                    R = right align
        cell_width: 
                    even: evenly distribute cell/column width
                    uneven: base cell size on lenght of cell/column items
                    int: int value for width of each cell/column
                    list of ints: list equal to number of columns with the widht of each cell / column
        x_start: 
                    where the left edge of table should start
        emphasize_data:  
                    which data elements are to be emphasized - pass as list 
        emphasize_style: the font style you want emphaized data to take
        emphasize_color: emphasize color (if other than black) 
        
        """
        default_style = self.font_style
        if emphasize_style == None:
            emphasize_style = default_style
        # default_font = self.font_family
        # default_size = self.font_size_pt
        # default_style = self.font_style
        # default_color = self.color # This does not work

        # Get Width of Columns
        def get_col_widths():
            col_width = cell_width
            if col_width == 'even':
                col_width = self.epw / len(data[0]) - 1  # distribute content evenly   # epw = effective page width (width of page not including margins)
            elif col_width == 'uneven':
                col_widths = []

                # searching through columns for largest sized cell (not rows but cols)
                for col in range(len(table_data[0])): # for every row
                    longest = 0 
                    for row in range(len(table_data)):
                        cell_value = str(table_data[row][col])
                        value_length = self.get_string_width(cell_value)
                        if value_length > longest:
                            longest = value_length
                    col_widths.append(longest + 4) # add 4 for padding
                col_width = col_widths



                        ### compare columns 

            elif isinstance(cell_width, list):
                col_width = cell_width  # TODO: convert all items in list to int        
            else:
                # TODO: Add try catch
                col_width = int(col_width)
            return col_width

        # Convert dict to lol
        # Why? because i built it with lol first and added dict func after
        # Is there performance differences?
        if isinstance(table_data, dict):
            header = [key for key in table_data]
            data = []
            for key in table_data:
                value = table_data[key]
                data.append(value)
            # need to zip so data is in correct format (first, second, third --> not first, first, first)
            data = [list(a) for a in zip(*data)]

        else:
            header = table_data[0]
            data = table_data[1:]

        line_height = self.font_size * 2.5

        col_width = get_col_widths()
        self.set_font(size=title_size)

        # Get starting position of x
        # Determin width of table to get x starting point for centred table
        if x_start == 'C':
            table_width = 0
            if isinstance(col_width, list):
                for width in col_width:
                    table_width += width
            else: # need to multiply cell width by number of cells to get table width 
                table_width = col_width * len(table_data[0])
            # Get x start by subtracting table width from pdf width and divide by 2 (margins)
            margin_width = self.w - table_width
            # TODO: Check if table_width is larger than pdf width

            center_table = margin_width / 2 # only want width of left margin not both
            x_start = center_table
            self.set_x(x_start)
        elif isinstance(x_start, int):
            self.set_x(x_start)
        elif x_start == 'x_default':
            x_start = self.set_x(self.l_margin)


        # TABLE CREATION #

        # add title
        if title != '':
            self.multi_cell(0, line_height, title, border=0, align='j', ln=3, max_line_height=self.font_size)
            self.ln(line_height) # move cursor back to the left margin

        self.set_font(size=data_size)
        # add header
        y1 = self.get_y()
        if x_start:
            x_left = x_start
        else:
            x_left = self.get_x()
        x_right = self.epw + x_left
        if  not isinstance(col_width, list):
            if x_start:
                self.set_x(x_start)
            for datum in header:
                self.set_font(style = emphasize_style)  # NEW CODE
                self.multi_cell(col_width, line_height, datum, border=0, align=align_header, ln=3, max_line_height=self.font_size)
                self.set_font(style = default_style)    # NEW CODE
                x_right = self.get_x()
            self.ln(line_height) # move cursor back to the left margin
            y2 = self.get_y()
            self.line(x_left,y1,x_right,y1)
            self.line(x_left,y2,x_right,y2)

            for row in data:
                if x_start: # not sure if I need this
                    self.set_x(x_start)
                for datum in row:
                    if datum in emphasize_data:
                        self.set_text_color(*emphasize_color)
                        self.set_font(style=emphasize_style)
                        self.multi_cell(col_width, line_height, datum, border=0, align=align_data, ln=3, max_line_height=self.font_size)
                        self.set_text_color(0,0,0)
                        self.set_font(style=default_style)
                    else:
                        self.multi_cell(col_width, line_height, datum, border=0, align=align_data, ln=3, max_line_height=self.font_size) # ln = 3 - move cursor to right with same vertical offset # this uses an object named self
                self.ln(line_height) # move cursor back to the left margin
        
        else:
            if x_start:
                self.set_x(x_start)
            for i in range(len(header)):
                datum = header[i]
                self.set_font(style = emphasize_style)  # NEW CODE
                self.multi_cell(col_width[i], line_height, datum, border=0, align=align_header, ln=3, max_line_height=self.font_size)
                self.set_font(style = default_style) # NEW CODE
                x_right = self.get_x()
            self.ln(line_height) # move cursor back to the left margin
            y2 = self.get_y()
            self.line(x_left,y1,x_right,y1)
            self.line(x_left,y2,x_right,y2)


            for i in range(len(data)):
                if x_start:
                    self.set_x(x_start)
                row = data[i]
                for i in range(len(row)):
                    datum = row[i]
                    if not isinstance(datum, str):
                        datum = str(datum)
                    adjusted_col_width = col_width[i]
                    if datum in emphasize_data:
                        self.set_text_color(*emphasize_color)
                        self.set_font(style=emphasize_style)
                        self.multi_cell(adjusted_col_width, line_height, datum, border=0, align=align_data, ln=3, max_line_height=self.font_size)
                        self.set_text_color(0,0,0)
                        self.set_font(style=default_style)
                    else:
                        self.multi_cell(adjusted_col_width, line_height, datum, border=0, align=align_data, ln=3, max_line_height=self.font_size) # ln = 3 - move cursor to right with same vertical offset # this uses an object named self
                self.ln(line_height) # move cursor back to the left margin
        y3 = self.get_y()
        self.line(x_left,y3,x_right,y3)

    school_list = {
        'James Ruse Agricultural High School' : 106,
        'Baulkham Hills High School' : 98, 
        'North Sydney Boys High School' : 93, 
        'Sydney Boys High School' : 92, 
        'North Sydney Girls High School' : 92, 
        'Nomanhurst Boys High School' : 92, 
        'Hornsby Girls High School' : 92, 
        'Girraween High School' : 92,
        'Sydney Girls High School' : 91, 
        'Fort Street High School' : 90, 
        'Chatswood High School (P)' : 90,
        'Penrith High School' : 87,
        'Northern Beaches Secondary College - Manly Campus' : 86,
        'Hurlstone Agricultural High School' : 86,
        'Sydney Technical High School' : 85, 
        'St George Girls High School' : 85, 
        'Ryde Secondary College (P)' : 85, 
        'Parramatta High School (P)' : 84, 
        'Caringbah High School' : 82, 
        'Blacktown Girls High School (P)' : 82,
        'Blacktown Boys High School (P)' : 82, 
        'Tempe High School (P)' : 80,
        'Sefton High School (P)' : 80, 
        'Merewether High School' : 80, 
        'Leichhardt Campus - Sydney Secondary College (P)' : 80, 
        'Gosford High School' : 80,
        'Smiths Hill High School' : 78, 
        'Richmond High School - Richmond Agricultural College (P)' : 78, 
        'Prairiewood High School (P)' : 77, 
        'Balmain Campus - Sydney Secondary College (P)' : 77,
        'Macquarie Fields High School (P)' : 76, 
        'Alexandria Park Community School (P)' : 75, 
        'Rose Bay Secondary College (P)' : 73, 
        'Moorebank High School (P)' : 73, 
        'Bonnyrigg High School (P)' : 72, 
        'Elizabeth Macarthur High School (P)' : 70, 
        'Auburn Girls High School (P)' : 70, 
        'Armidale Secondary College (P)' : 70, 
        'Peel High School (P)' : 65, 
        'Kooringal High School (P)' : 65, 
        'Karabar High School (P)' : 65, 
        'Granville Boys High School (P)' : 65, 
        'Grafton High School (P)' : 65, 
        'Gorokan High School (P)' : 65
    }
    
    def school_table(self, mark):
        self.set_font('helvetica', 'B', 11)
        self.cell(50, 1.25*self.font_size, 'Rank', border = 1)
        self.cell(130, 1.25*self.font_size, 'School', border = 1)
        self.ln(1.25*self.font_size)
        self.set_font('helvetica', '', 10)
        for i in range(0, len(PDF.school_list)):
            for j in range(0, 2):
                if j == 0:
                    self.cell(50, 1.25*self.font_size, str(i+1), border = 1, align = 'C')
                    
                if j == 1: 
                    # change colour 
                    if mark + 2 < list(PDF.school_list.values())[i]:
                        self.set_text_color(255, 0, 0)
                    elif mark - 2 < list(PDF.school_list.values())[i]:
                        self.set_text_color(255, 153, 51)
                    else: 
                        self.set_text_color(0, 204, 102)
                    
                    self.cell(130, 1.25*self.font_size, list(PDF.school_list.keys())[i], border = 1, align = 'C')
                    self.set_text_color(0, 0, 0)   
            self.ln(1.25*self.font_size)  
        
        self.ln(5)   
        
    school_list_perc = {
        'James Ruse Agricultural High School' : 0.075,
        'Baulkham Hills High School' : 0.15, 
        'North Sydney Boys High School' : 0.15, 
        'Sydney Boys High School' : 0.2, 
        'North Sydney Girls High School' : 0.2, 
        'Nomanhurst Boys High School' : 0.2, 
        'Sydney Girls High School' : 0.25, 
        'Hornsby Girls High School' : 0.275, 
        'Fort Street High School' : 0.3,
        'Girraween High School' : 0.325,
        'Chatswood High School (P)' : 0.35,
        'Penrith High School' : 0.35,
        'Northern Beaches Secondary College - Manly Campus' : 0.4,
        'Hurlstone Agricultural High School' : 0.4,
        'Parramatta High School (P)' : 0.4, 
        'Sydney Technical High School' : 0.45, 
        'St George Girls High School' : 0.45, 
        'Ryde Secondary College (P)' : 0.45, 
        'Caringbah High School' : 0.45, 
        'Blacktown Girls High School (P)' : 0.6,
        'Blacktown Boys High School (P)' : 0.6, 
        'Tempe High School (P)' : 0.7,
        'Sefton High School (P)' : 0.7, 
        'Merewether High School' : 0.7, 
        'Leichhardt Campus - Sydney Secondary College (P)' : 0.7, 
        'Gosford High School' : 0.7,
        'Smiths Hill High School' : 0.75, 
        'Richmond High School - Richmond Agricultural College (P)' : 0.75, 
        'Prairiewood High School (P)' : 0.8, 
        'Balmain Campus - Sydney Secondary College (P)' : 0.8,
        'Macquarie Fields High School (P)' : 0.8, 
        'Alexandria Park Community School (P)' : 0.8, 
        'Rose Bay Secondary College (P)' : 0.85, 
        'Moorebank High School (P)' : 0.85, 
        'Bonnyrigg High School (P)' : 0.85, 
        'Elizabeth Macarthur High School (P)' : 0.9, 
        'Auburn Girls High School (P)' : 0.9, 
        'Armidale Secondary College (P)' : 0.9, 
        'Peel High School (P)' : 1, 
        'Kooringal High School (P)' : 1, 
        'Karabar High School (P)' : 1, 
        'Granville Boys High School (P)' : 1, 
        'Grafton High School (P)' : 1,
        'Gorokan High School (P)' : 1
    }
    
    school_list_perc_oc = {
        'Beecroft Public School' : 0.05, 
        'North Rocks Public School' : 0.1, 
        'Dural Public School' : 0.1, 
        'Ermington Public School' : 0.125, 
        'Matthew Pearce Public School' : 0.125, 
        'Summer Hill Public School' : 0.125, 
        'Waitara Public School' : 0.15, 
        'Chatswood Public School' : 0.15, 
        'Artarmon Public School' : 0.15, 
        'Ryde Public School' : 0.175, 
        'Ironbark Ridge Public School' : 0.2, 
        'Ashfield Public School' : 0.225, 
        'Greystanes Public School' : 0.225, 
        'Quakers Hill Public School' : 0.25, 
        'Georges Hall Public School' : 0.25, 
        'Neutral Bay Public School' : 0.25, 
        'Hurstville Public School' : 0.25, 
        'Holsworthy Public School' : 0.3, 
        'Woollahra Public School' : 0.3, 
        'Blacktown South Public School' : 0.3, 
        'Picnic Point Public School' : 0.3, 
        'Balmain Public School' : 0.35, 
        'Newbridge Heights Public School' : 0.35, 
        'Earlwood Public School' : 0.35, 
        'Balgowlah Heights Public School' : 0.35, 
        'Leumeah Public School' : 0.35, 
        'Sutherland Public School' : 0.35, 
        'Wilkins Public School' : 0.35, 
        'Caringbah North Public School' : 0.4,
        'Wollongong Public School' : 0.4, 
        'Mona Vale Public School' : 0.4, 
        'Blaxcell Street Public School' : 0.4, 
        'Casula Public School' : 0.4, 
        'Harrington Street Public School' : 0.45, 
        'Smithfield Public School' : 0.45, 
        'New Lambton South Public School' : 0.45, 
        'Alexandria Park Community School' : 0.45, 
        'St Johns Park Public School' : 0.45, 
        'Greenacre Public School' : 0.5, 
        'Gosford Public School' : 0.5, 
        'St Andrews Public School' : 0.55, 
        'Kingswood Public School' : 0.55, 
        'Bradbury Public School' : 0.6, 
        'Camden South Public School' : 0.6, 
        'Colyton Public School' : 0.6, 
        'Tighes Hill Public School' : 0.6, 
        'Glenbrook Public School' : 0.6, 
        'Rutherford Public School' : 0.7, 
        'Soldiers Point Public School' : 0.7, 
        'Tamworth Public School' : 0.7, 
        'Jewells Primary School' : 0.75, 
        'Armidale City Public School' : 0.75, 
        'Maryland Public School' : 0.75, 
        'Richmond Public School' : 0.75, 
        'Queanbeyan South Public School' : 0.75, 
        'Wyong Public School' : 0.8, 
        'Wentworth Falls Public School' : 0.8, 
        'Cessnock West Public School' : 0.8, 
        'Sturt Public School' : 0.8, 
        'Coonabarabran Public School' : 0.8, 
        'Cudgegong Valley Public School' : 0.8, 
        'Biraban Public School' : 0.85, 
        'Goulburn West Public School' : 0.85, 
        'Moree Public School' : 0.85, 
        'Toormina Public School' : 0.85, 
        'Alstonville Public School' : 0.9, 
        'Dubbo West Public School' : 0.9, 
        'Illaroo Road Public School' : 0.9, 
        'Lithgow Public School': 0.9,
        'South Grafton Public School' : 0.9, 
        'Tahmoor Public School' : 0.9, 
        'Tamworth South Public School': 0.9, 
        'Bathurst West Public School':0.9, 
        'Goonellabah Public School': 0.9, 
        'Port Macquarie Public School': 0.9 
    }
    
    def school_table_perc(self, percentile):
        self.set_font('helvetica', 'B', 11)
        self.cell(50, 1.25*self.font_size, 'Rank', border = 1)
        self.cell(130, 1.25*self.font_size, 'School', border = 1)
        self.ln(1.25*self.font_size)
        self.set_font('helvetica', '', 10)
        for i in range(0, len(PDF.school_list_perc)):
            for j in range(0, 2):
                if j == 0:
                    self.cell(50, 1.25*self.font_size, str(i+1), border = 1, align = 'C')
                    
                if j == 1: 
                    # change colour 
                    if percentile - 0.03 > list(PDF.school_list_perc.values())[i]:
                        self.set_text_color(255, 0, 0)
                    elif percentile + 0.03 > list(PDF.school_list_perc.values())[i]:
                        self.set_text_color(255, 153, 51)
                    else: 
                        self.set_text_color(0, 204, 102)
                    
                    self.cell(130, 1.25*self.font_size, list(PDF.school_list_perc.keys())[i], border = 1, align = 'C')
                    self.set_text_color(0, 0, 0)   
            self.ln(1.25*self.font_size)  
        
        self.ln(5)   
    
    def school_table_perc_oc(self, percentile):
        self.set_font('helvetica', 'B', 11)
        self.cell(50, 1.25*self.font_size, 'Rank', border = 1)
        self.cell(130, 1.25*self.font_size, 'School', border = 1)
        self.ln(1.25*self.font_size)
        self.set_font('helvetica', '', 10)
        for i in range(0, len(PDF.school_list_perc_oc)):
            for j in range(0, 2):
                if j == 0:
                    self.cell(50, 1.25*self.font_size, str(i+1), border = 1, align = 'C')
                    
                if j == 1: 
                    # change colour 
                    if percentile - 0.03 > list(PDF.school_list_perc_oc.values())[i]:
                        self.set_text_color(255, 0, 0)
                    elif percentile + 0.03 > list(PDF.school_list_perc_oc.values())[i]:
                        self.set_text_color(255, 153, 51)
                    else: 
                        self.set_text_color(0, 204, 102)
                    
                    self.cell(130, 1.25*self.font_size, list(PDF.school_list_perc_oc.keys())[i], border = 1, align = 'C')
                    self.set_text_color(0, 0, 0)   
            self.ln(1.25*self.font_size)  
        
        self.ln(5)   
        
    
    def create_mark_table(self, sub_table):
        self.create_table(table_data = sub_table, data_size = 14, cell_width = [150, 30], 
                 emphasize_data = ['Your Mark:'], emphasize_style = 'B')
    
    
    def questions_table(self, q_file, ans_sheet):
        try:
            questions = pd.read_excel(q_file)
        except:
            questions = pd.read_csv(q_file)
        q_table = [['Question', 'Type', 'Correct Answer', '% Students Correct']]
        q_table2 = [[str(i+1), questions.iloc[i, 0], questions.iloc[i, 1], '{:.2f}%'.format(np.sum(ans_sheet.iloc[:, i])/np.sum(~pd.isna(ans_sheet.iloc[:, i]))*100)] for i in range(0, len(questions))]
        return(q_table + q_table2) 
    
    
    # table without answers
    # def create_question_table(self, q_table):
    #     self.set_font('cambria', '', 12)
    #     for row in q_table:
    #         counter = 0
    #         for data in row:
    #             if counter == 0:
    #                 self.cell(20, 1.25*self.font_size, data, border = 1)
    #             if counter == 1:
    #                 self.cell(120, 1.25*self.font_size, data, border = 1)
    #             if counter == 2:
    #                 self.cell(40, 1.25*self.font_size, data, border = 1, align = 'C')
    #             counter += 1
    #         self.ln(1.25*self.font_size)
    
    # table with answers
    def create_question_table(self, q_table, ans_sheet, did_test):
        self.set_font('cambria', '', 12)
        qnum = -1
        for row in q_table:
            counter = 0
            for data in row:
                if counter == 0:
                    self.cell(20, 1.25*self.font_size, data, border = 1)
                if counter == 1:
                    if did_test:
                        if ans_sheet[qnum] == 0 and qnum != -1:
                            self.set_text_color(255, 0, 0)
                        elif qnum != -1:
                            self.set_text_color(0, 204, 102)
                        self.cell(90, 1.25*self.font_size, data, border = 1)
                        self.set_text_color(0, 0, 0)
                    else:
                        self.cell(90, 1.25*self.font_size, data, border = 1)
                if counter == 2:
                    self.cell(30, 1.25*self.font_size, data, border = 1, align = 'C')
                if counter == 3:
                    self.cell(40, 1.25*self.font_size, data, border = 1, align = 'C')
                counter += 1
            qnum += 1
            self.ln(1.25*self.font_size)


    