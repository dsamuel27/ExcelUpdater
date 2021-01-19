import pandas as pd
import PySimpleGUI as sg
import os

def can_update(source,master):
    if(source == master or source == '' or master == ''):
        return False

    try:
        src_df = pd.read_excel(source)
        print(src_df.columns)
        master_df = pd.read_excel(master)
        print(master_df.columns)
        equality = src_df.columns == master_df.columns
        return not False in equality
    except ValueError:
        return False

def launch_update(source,master):
    src_df = pd.read_excel(source)
    master_df = pd.read_excel(master)
    cols = src_df.columns.tolist()
    print(src_df)
    row1 = [sg.Text('Select column you want to match on')]
    row2 = [sg.Combo(cols,key='Col')]
    btn = [sg.Button('Update')]
    window = sg.Window(title='Excel Updater', layout=[row1, row2,btn], margins=(100, 50))

    while True:
        event, values = window.Read()
        if event in (None, 'Exit'):
            break
        if event == 'Update':
            col = values['Col']
            print(col)
            if (col != ''):
                src_df.to_excel('source_backup.xls')
                master_df.to_excel('master_backup.xls')

                all_cols_but_org = master_df.columns.tolist()
                all_cols_but_org.remove(col)
                print(all_cols_but_org)
                master_df.loc[master_df[col].isin(src_df[col]), all_cols_but_org] = src_df[all_cols_but_org]
                master_df.set_index(col,inplace=True)
                master_df.to_excel(master)
                break;

    window.Close()

def main():

    row1 = [sg.Text('Select Excel file from which contains update data')]

    xls_in_dir = []
    for root,dirs ,files in os.walk(os.getcwd()):
        for name in files:
            ext = name.split('.')
            if(len(ext) == 2):
                if ext[1] == 'xls':
                    xls_in_dir.append(name)

    row2 = [sg.Combo(xls_in_dir,key='update_key')]
    row3 = [sg.Text('Select Excel file to be updated')]
    row4 = [sg.Combo(xls_in_dir,key='master_key')]

    submit_btn = [sg.Button('Update')]

    window = sg.Window(title='Excel Updater',layout=[row1,row2,row3,row4,submit_btn], margins=(100,50))

    while True:
        event, values = window.Read()
        if event in (None, 'Exit'):
            break
        if event == 'Update':
            source = values['update_key']
            master = values['master_key']
            if(can_update(source,master)):
                launch_update(source,master)
                break;
            else:
                print('Invalid')
    window.Close()




if __name__ == '__main__':
    main()