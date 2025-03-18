# Statistics Plots Generation Module
import seaborn as sns
import matplotlib.pyplot as plt
import numpy as np

sns.set_theme(palette = 'Set2')
colors = ['#435F8E', '#282A40']
sns.set_palette(sns.color_palette(colors))

# maths
def maths_chart(data, student_mark):
    plt.ioff()
    fig = plt.figure(figsize=[6.4, 4])
    sns.histplot(data.maths_mark[~np.isnan(data.maths_mark.astype(np.float64))], kde = True, discrete = True)
    plt.ylabel('Number of Students')
    plt.xlabel('Mark')
    plt.title('Mathematical Reasoning')
    plt.rc('xtick', labelsize = 8)
    plt.rc('ytick', labelsize = 8)
    plt.rc('axes', labelsize = 10)
    if student_mark >= 0:
        plt.axvline(x = student_mark, color = '#BE413C', label = 'Student Mark')
    plt.axvline(x = np.mean(data.maths_mark), color = '#F0AF41', label = 'Cohort Mean')
    plt.legend()
    return(fig)

# reading
def reading_chart(data, student_mark):
    plt.ioff()
    fig = plt.figure(figsize=[6.4, 4])
    sns.histplot(data.reading_mark[~np.isnan(data.reading_mark.astype(np.float64))], kde = True, discrete = True)
    plt.ylabel('Number of Students')
    plt.xlabel('Mark')
    plt.title('Reading')
    plt.rc('xtick', labelsize = 8)
    plt.rc('ytick', labelsize = 8)
    plt.rc('axes', labelsize = 10)
    if student_mark >= 0:
        plt.axvline(x = student_mark, color = '#BE413C', label = 'Student Mark')
    plt.axvline(x = np.mean(data.reading_mark), color = '#F0AF41', label = 'Cohort Mean')
    plt.legend()
    return(fig)

# thinking skills
def thinking_chart(data, student_mark): 
    plt.ioff()
    fig = plt.figure(figsize=[6.4, 4])
    sns.histplot(data.thinking_mark[~np.isnan(data.thinking_mark.astype(np.float64))], kde = True, discrete = True)
    plt.ylabel('Number of Students')
    plt.xlabel('Mark')
    plt.title('Thinking Skills')
    plt.rc('xtick', labelsize = 8)
    plt.rc('ytick', labelsize = 8)
    plt.rc('axes', labelsize = 10)
    if student_mark >= 0:
        plt.axvline(x = student_mark, color = '#BE413C', label = 'Student Mark')
    plt.axvline(x = np.mean(data.thinking_mark), color = '#F0AF41', label = 'Cohort Mean')
    plt.legend()
    return(fig)

# writing
def writing_chart(data, student_mark):
    plt.ioff()
    fig = plt.figure(figsize=[6.4, 4])
    sns.histplot(data.writing_mark[~np.isnan(data.writing_mark.astype(np.float64))], kde = True, discrete = True)
    plt.ylabel('Number of Students')
    plt.xlabel('Mark')
    plt.title('Writing')
    plt.rc('xtick', labelsize = 8)
    plt.rc('ytick', labelsize = 8)
    plt.rc('axes', labelsize = 10)
    if student_mark >= 0:
        plt.axvline(x = student_mark, color = '#BE413C', label = 'Student Mark')
    plt.axvline(x = np.mean(data.writing_mark), color = '#F0AF41', label = 'Cohort Mean')
    plt.legend()
    return(fig)