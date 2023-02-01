                                                                                                                                                                                              #plan1: use pandas to load the file directly
import pandas
import matplotlib.pyplot as plt
import os
import win32com.client as win32
import pandas as pd
import numpy as np
from pathlib import Path
from pywintypes import com_error


def draw_bar_chart(df, title, output_folder):                                                                                                                                     
    '''
   Draw a stacked bar chart: The first element with the green color, the second element with the yellow color
    Extra highlight: let the maximum value among the second elements to red color
    '''
    # find the index location of the maximum value of the second element
    # like 0, 1, 2...
    max_index = df.iloc[:, 1].idxmax()
    max_index_location = df.index.get_loc(max_index)
    fig, ax = plt.subplots()
    # enlarge the figure size according to length of the index
    fig.set_size_inches(2 * len(df.index), 5)
    # set the label by the column name
    ax.bar(range(len(df.index)), df.iloc[:, 0], color='green', label = df.columns[0])
    # check whether any value of the second element is larger than 0
    if df.iloc[max_index_location, 1] > 0:
        ax.bar(max_index_location, df.iloc[max_index_location, 1], bottom=df.iloc[max_index_location, 0], color='red', label= df.columns[1]+" (maximum value)")
    if df.iloc[:max_index_location, 1].max() > 0 or df.iloc[max_index_location+1:, 1].max() > 0:
        ax.bar(range(max_index_location), df.iloc[:max_index_location, 1], bottom=df.iloc[:max_index_location, 0], color='yellow', label= df.columns[1])
        ax.bar(range(max_index_location+1, len(df.index)), df.iloc[max_index_location+1:, 1], bottom=df.iloc[max_index_location+1:, 0], color='yellow')
        
    # show the label at the right top of the bar
    ax.legend(loc='upper right')
    # ax.bar(range(len(df.index)-1), df.iloc[:, 1], bottom=df.iloc[:, 0], color='yellow')
    ax.bar
    # set the x-axis by the index
    ax.set_xticks(range(len(df.index)))
    ax.set_xticklabels(df.index)
    
    for i in range(len(df)):
        if df.iloc[i, 0] == 0:
            plt.text(i, df.iloc[i, 1]/2, int(df.iloc[i, 1]), ha='center', va='center')
        elif df.iloc[i, 1] == 0:
            plt.text(i, df.iloc[i, 0]/2, int(df.iloc[i, 0]), ha='center', va='center')
        else:
            plt.text(i, df.iloc[i, 0] / 2,  int(df.iloc[i, 0]), ha='center', va='center')
            plt.text(i, (2 * df.iloc[i, 0] + df.iloc[i,1]) / 2,  int(df.iloc[i, 1]), ha='center', va='center')
    plt.title(title)
    # save the image to the output folder
    plt.savefig(os.path.join(output_folder, title + '.png'))



def extract_gap(input_folder):
    ''' 
    Extract the gap between the last two columns
    for the status of the same file name, it should be counted as a whole
    '''
    status_data = dict()
    file_number = dict()
    files = os.listdir(input_folder)
    for file in files:
        file_name = file.split('.')[0].split('_')[0]
        file_path = os.path.join(input_folder, file)
        # read the sixth sheet (pivot_table) and set the six row as the column name
        df = pandas.read_excel(file_path, sheet_name=5, header=5)
        # check whether there is a column named "None"
        if 'None' not in df.columns:
            df['None'] = 0
            df = df[['None', df.columns[-2]]]
        else:
            df = df[['None', df.columns[-1]]]
        df = df.iloc[-1:]
        # "NN" is sum (last element) minus "None"
        df['NN'] = df[df.columns[-1]] - df['None']
        # drop the the total column
        df = df.drop(df.columns[-2], axis=1)
        # set all the index as 0
        df.index = [0]
        # add the dataframe to the dictionary
        # if the file name is the same, add the dataframe value to the same key
        if file_name in status_data.keys():
            status_data[file_name] = status_data[file_name] + df
            file_number[file_name] += 1

        else:
            status_data[file_name] = df
            file_number[file_name] = 1
    # convert the dictionary to dataframe, the key is the index
    df = pandas.concat(status_data)
    # rename the column name
    df = df.rename(columns={'None': 'Number of gap', 'NN': 'Number of fixed'})
    # exchange the two elements in df 
    df = df[['Number of fixed', 'Number of gap']]
    # set the index by the first element of the index
    df.index = df.index.map(lambda x: x[0])
    # rename the index name by adding the file_number
    df.index = df.index.map(lambda x: (x, file_number[x]))
    return df



def extract_linked_request(input_folder):
    '''
    use win32 read the eighth sheet (pivot_table)
    '''
    # give the absolute path of the files
    linked_request_data = dict()
    file_number = dict()
    files = os.listdir(input_folder)
    for file in files:
        file_name = file.split('.')[0].split('\\')[-1].split('_')[0]
        file_path = os.path.join(input_folder, file)
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        # excel.Visible = True
        try:
            wb = excel.Workbooks.Open(file_path)
        except com_error as e:
            if e.excepinfo[5] == -2146827284:
                print(f'Failed to open spreadsheet.  Invalid filename or location: {file_path}')
                exit()
            else:
                raise e
        ws = wb.Worksheets(8)
        # select the range of the pivot table
        pvtTable = ws.Range('C3').PivotTable
        page_range_item = []
        for i in pvtTable.PageRange:
            page_range_item.append(str(i))
        # set filter as "Agreed"
        pvtTable.PivotFields(page_range_item[0]).CurrentPage = 'Agreed'
        data = pvtTable.DataBodyRange.Value
        # extract the first element and the last element of the list
        df = pandas.DataFrame(data[-1][0:]).T
        # just rename the last column as 'Sum' and the first column as 'Number of required link'
        # do not change other column names
        df.rename(columns={df.columns[-1]: 'Sum', df.columns[0]: 'Number of required link'}, inplace=True)
        # and then create another element by last column minus the first column
        df['Number of requirednotlink'] = df['Sum'] - df['Number of required link']
        # keep the first and the last column
        df = df[['Number of required link', 'Number of requirednotlink']]
        if file_name in linked_request_data.keys():
            linked_request_data[file_name] = linked_request_data[file_name] + df
            file_number[file_name] += 1
        else:
            linked_request_data[file_name] = df
            file_number[file_name] = 1
        wb.Close(False)

    excel.Application.Quit()  
    # convert the dict to dataframe, the key is the index
    df = pandas.concat(linked_request_data)
    # set the first column "Number of requiredlink", the second column "Number of requirednotlink"
    # df.columns = ['Number of required link', 'Number of requirednotlink']
    # set the index by the first element of the index
    df.index = df.index.map(lambda x: x[0])
    df.index = df.index.map(lambda x: (x, file_number[x]))
    return df


def auto_draw(input_folder, output_folder, document_type):
    df_gap = extract_gap(input_folder)
    df_link = extract_linked_request(input_folder)
    # draw the bar chart with document_type
    draw_bar_chart(df_gap, f'Project {document_type} Report Status', output_folder)
    draw_bar_chart(df_link, f'Project {document_type} Report Link', output_folder)



if __name__ == '__main__':
    # give the absolute path of the folder
    input_folder = r'C:\Users\sophie\OneDrive\桌面\autoDraw\example'
    output_folder = r'C:\Users\sophie\OneDrive\桌面\autoDraw\output'
    if not os.path.exists(output_folder):
        os.mkdir(output_folder)
    df_gap = extract_gap(input_folder)
    df_link = extract_linked_request(input_folder)
    draw_bar_chart(df_gap,'Gap condition for every project about status in certain time period', output_folder)
    draw_bar_chart(df_link, 'Link condition for every project about inlink in certain time period', output_folder)
