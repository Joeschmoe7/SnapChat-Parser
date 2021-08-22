import sys
import PySimpleGUI as sg
import zipfile
import pandas as pd
from pandas import DataFrame
import os
import numpy as np

def inputfile():  #choose file dialogue box
    global fname
    if len(sys.argv) == 1:
        fname = sg.popup_get_file('Choose a SnapChat ZIP file.',
                                      title=None,
                                      default_path="",
                                      default_extension="",
                                      save_as=False,
                                      multiple_files=False,
                                      file_types=(('ZIP', '.zip'),),
                                      no_window=False,
                                      size=(None, None),
                                      button_color=None,
                                      background_color=None,
                                      text_color=None,
                                      icon=None,
                                      font=None,
                                      no_titlebar=False,
                                      grab_anywhere=False,
                                      keep_on_top=False,
                                      location=(None, None),
                                      initial_folder=None)
    else:
        fname = sys.argv[1]
    
    if not fname:
        sg.popup("You didn't pick a file")
        raise SystemExit("Cancelling: no filename supplied")
    else:
        sg.popup('Successfully added!')
inputfile()
def fileexist(): # Success pop up if new file exists.
    
    if os.path.exists(savefname) or os.path.exists(savefname + ".xlsx"):
        sg.popup("Success!")
    else:
        sg.popup('Uh oh!  Something went wrong.') 

z = zipfile.ZipFile(fname, "r")  #open ZIP file
fileList = []

for filenames in z.namelist():
    fileList = z.namelist()
    
df = pd.DataFrame() #declare DataFrame so it can be used to test for null
df2 = pd.DataFrame()
df3 = pd.DataFrame()
df4 = pd.DataFrame()
df5 = pd.DataFrame()
df6 = pd.DataFrame()
df7 = pd.DataFrame()
df8 = pd.DataFrame()
df9 = pd.DataFrame()


def saveBox():  #open Save file dialogue box
    global savefname
    layout = [[sg.Text('Amost done!  Enter the file name full path for the new file.')],
              [sg.Text('Save to:', size=(8, 1)), sg.Input(), sg.SaveAs(file_types=(("XLSX", ".xlsx"),))],
              [sg.Submit(), sg.Cancel()]]
        
    window = sg.Window('Save file location:', layout)
        
    event, values = window.read()
    window.close()
    
    savefname = values[0]
    if ".xlsx" not in (savefname):
        savefname = savefname + ".xlsx"

#open Subsciber_information.csv

if "subscriber_information.csv" in fileList:
    csvfile3 = pd.read_csv(z.open("subscriber_information.csv"), encoding = "utf-8", engine = "python", skiprows = 3, parse_dates= ["created"])            
    csvfile3['created'] = csvfile3['created'] .apply(lambda x: pd.Timestamp(x) .strftime('%b %d, %Y %H:%M:%S  ') if x == x else np.nan)
    df3 = DataFrame(csvfile3, columns=['id', 'email_address', 'created', 'creation_ip', 'phone_number', 'display_name', 'status'])        
    df3['created'] = pd.to_datetime(df3["created"])  #recognize column as date
    df3.sort_values(by=['created'], inplace=True)  #sort by date
    df3 = df3.drop_duplicates()  #delete deuplcate rows
    df3.columns = ['ID', 'Email Address', 'Creation Date (UTC)', 'Creation IP', 'Phone Number', 'Display Name', 'Account Status']
    df3 = df3.fillna('') #replace null values
    df3 = df3.applymap(str)  #converts dataframe to a string for better formatting.
    
#open ip_data.csv
if "ip_data.csv" in fileList:
    csvfile4 = pd.read_csv(z.open("ip_data.csv"), encoding = "utf-8", engine = "python", skiprows = 3, parse_dates= ["timestamp"])            
    csvfile4['timestamp'] = csvfile4['timestamp'] .apply(lambda x: pd.Timestamp(x) .strftime('%b %d, %Y %H:%M:%S  ') if x == x else np.nan)
    df4 = DataFrame(csvfile4, columns=['timestamp', 'type', 'ip'])        
    df4['timestamp'] = pd.to_datetime(df4["timestamp"])
    df4.sort_values(by=['timestamp'], inplace=True)
    df4 = df4.drop_duplicates()
    df4.columns = ['Date (UTC)', 'Type', 'IP Address']
    df4 = df4.fillna('')
    df4 = df4.applymap(str)

#open Chat.CSV file          
if "chat.csv" in fileList:
    csvfile = pd.read_csv(z.open("chat.csv"), encoding = "utf-8", engine = "python", skiprows = 3, parse_dates= ["timestamp"])            
    csvfile['timestamp'] = csvfile['timestamp'] .apply(lambda x: pd.Timestamp(x) .strftime('%b %d, %Y %H:%M:%S  ') if x == x else np.nan)
    df = DataFrame(csvfile, columns=['timestamp', 'id', 'from', 'to', 'body', 'href', 'media_id', 'saved'])        
    df['timestamp'] = pd.to_datetime(df["timestamp"])
    df.sort_values(by=['timestamp'], inplace=True)
    df = df.drop_duplicates()
    for files in fileList:
        for ind in df['media_id'].index:
            if ind in files:
                ind = files
    df['media_id'] = pd.DataFrame(links)
    df.columns = ['Date (UTC)', 'ID', 'From', 'To', 'Body', 'Href', 'Media-ID', 'Saved']
    df = df.fillna('')
    df = df.applymap(str)

        
#open Group Chat file
if "group-chat.csv" in fileList:
    csvfile2 = pd.read_csv(z.open("group-chat.csv"), encoding = "utf-8", engine = "python", skiprows = 3, parse_dates= ["timestamp"])            
    csvfile2['timestamp'] = csvfile2['timestamp'] .apply(lambda x: pd.Timestamp(x) .strftime('%b %d, %Y %H:%M:%S  ') if x == x else np.nan)
    df2 = DataFrame(csvfile2, columns=['timestamp', 'message_type', 'from', 'to_group_name', 'to_group_id', 'text', 'href', 'media_id'])        
    df2['timestamp'] = pd.to_datetime(df2["timestamp"])
    df2.sort_values(by=['timestamp'], inplace=True)
    df2 = df2.drop_duplicates()
    df2.columns = ['Date (UTC)', 'Message Type', 'From', 'To Group Name', 'To Group ID', 'Body', 'Href', 'Media-ID']         
    df2 = df2.fillna('')
    df2 = df2.applymap(str)

if "friends_list.csv" in fileList:
    csvfile5 = pd.read_csv(z.open("friends_list.csv"), encoding = "utf-8", engine = "python", skiprows = 1)            
    df5 = DataFrame(csvfile5, columns=['Friend List'])        
    df5.columns = ["Friend's List"]
    df5 = df5.fillna('')
    df5 = df5.applymap(str)
    
if "snap.csv" in fileList:
    csvfile6 = pd.read_csv(z.open("snap.csv"), encoding = "utf-8", engine = "python", skiprows = 13, parse_dates= ["received", "viewed", "sent"])            
    csvfile6['received'] = csvfile6['received'] .apply(lambda x: pd.Timestamp(x) .strftime('%b %d, %Y %H:%M:%S  ') if x == x else np.nan) #if statement to ignore blank cells
    csvfile6['viewed'] = csvfile6['viewed'] .apply(lambda x: pd.Timestamp(x) .strftime('%b %d, %Y %H:%M:%S  ') if x == x else np.nan)
    csvfile6['sent'] = csvfile6['sent'] .apply(lambda x: pd.Timestamp(x) .strftime('%b %d, %Y %H:%M:%S  ') if x == x else np.nan)
    df6 = DataFrame(csvfile6, columns=['received', 'viewed', 'sent', 'id', 'media_type', 'sender', 'recipient', 'expired','timer', 'screenshotted', 'screenshot_count'])        
    df6['received'] = pd.to_datetime(df6["received"])
    df6.sort_values(by=['received'], inplace=True)
    df6 = df6.drop_duplicates()
    df6.columns = ['Date Received (UTC)', 'Date Viewed (UTC) ', 'Date Sent (UTC) ', 'ID', 'Media Type', 'Sender', 'Recipient', 'Expired', 'Timer', 'Screenshotted', 'Screenshot Count']
    df6 = df6.fillna('')
    df6 = df6.applymap(str)
    df6["Media Type"] = df6["Media Type"].replace('0', "Image")
    df6["Media Type"] = df6["Media Type"].replace('1', "Video with Sound")
    df6["Media Type"] = df6["Media Type"].replace('2', "Video without Sound")
    df6["Media Type"] = df6["Media Type"].replace('3', "Friend Request")

if "group_story_snap.csv" in fileList:
    csvfile7 = pd.read_csv(z.open("group_story_snap.csv"), encoding = "utf-8", engine = "python", skiprows = 3, parse_dates= ["post_timestamp"])            
    csvfile7['post_timestamp'] = csvfile7['post_timestamp'] .apply(lambda x: pd.Timestamp(x) .strftime('%b %d, %Y %H:%M:%S  '))
    df7 = DataFrame(csvfile7, columns=['post_timestamp', 'snap_id', 'group_story_id', 'current_group_name', 'media_type', 'media_length'])        
    df7.columns = ["Date/Time (UTC)", "Snap-ID", "Group Story ID", "Current Group Name", "Media Type", "Media Length"]
    df7 = df7.fillna('')
    df7 = df7.applymap(str)
    df7["Media Type"] = df7["Media Type"].replace('0', "Image")
    df7["Media Type"] = df7["Media Type"].replace('1', "Video with Sound")
    df7["Media Type"] = df7["Media Type"].replace('2', "Video without Sound")
    df7["Media Type"] = df7["Media Type"].replace('3', "Friend Request")
    
if "group.csv" in fileList:
    csvfile8 = pd.read_csv(z.open("group.csv"), encoding = "utf-8", engine = "python", skiprows = 3, parse_dates= ["latest_join_timestamp"])            
    csvfile8['latest_join_timestamp'] = csvfile8['latest_join_timestamp'] .apply(lambda x: pd.Timestamp(x) .strftime('%b %d, %Y %H:%M:%S  ') if x == x else np.nan)
    df8 = DataFrame(csvfile8, columns=['group_id', 'group_name', 'latest_join_timestamp'])        
    df8['latest_join_timestamp'] = pd.to_datetime(df8["latest_join_timestamp"])
    df8.sort_values(by=['latest_join_timestamp'], inplace=True)
    df8 = df8.drop_duplicates()
    df8.columns = ['Group ID', 'Group Name', 'Last Join Timestamp (UTC)']
    df8 = df8.fillna('')
    df8 = df8.applymap(str)
    
if "story.csv" in fileList:
    csvfile9 = pd.read_csv(z.open("story.csv"), encoding = "utf-8", engine = "python", skiprows = 3, parse_dates= ["timestamp"])            
    csvfile9['timestamp'] = csvfile9['timestamp'] .apply(lambda x: pd.Timestamp(x) .strftime('%b %d, %Y %H:%M:%S  ') if x == x else np.nan)
    df9 = DataFrame(csvfile9, columns=['timestamp', 'media_id', 'username', 'media_type', 'time', 'filter_id'])        
    df9['timestamp'] = pd.to_datetime(df9["timestamp"])
    df9.sort_values(by=['timestamp'], inplace=True)
    df9 = df9.drop_duplicates()
    df9.columns = ['Date/Time (UTC)', 'Media ID', 'Username', 'Media Type', 'Time', "Filter-ID"]
    df9 = df9.fillna('')
    df9 = df9.applymap(str)
    df9["Media Type"] = df9["Media Type"].replace('0', "Image")
    df9["Media Type"] = df9["Media Type"].replace('1', "Video with Sound")
    df9["Media Type"] = df9["Media Type"].replace('2', "Video without Sound")
    df9["Media Type"] = df9["Media Type"].replace('3', "Friend Request")
    
def saveXLSX():
    
    writer = pd.ExcelWriter(savefname, engine='xlsxwriter')
    workbook = writer.book
    
    if not df3.empty:
        df3.to_excel(writer, index=False, sheet_name="Subscriber Info")
        worksheet3 = writer.sheets['Subscriber Info']       
        worksheet3.set_column('A:A', 17)
        worksheet3.set_column('B:B', 30)
        worksheet3.set_column('C:C', 18)
        worksheet3.set_column('D:D', 37)
        worksheet3.set_column('E:E', 17)
        worksheet3.set_column('F:F', 20)
        worksheet3.set_column('G:G', 18)
     
    if not df4.empty:
        df4.to_excel(writer, index=False, sheet_name="IP History")
        worksheet4 = writer.sheets['IP History']       
        worksheet4.set_column('A:A', 19)
        worksheet4.set_column('B:B', 15)
        worksheet4.set_column('C:C', 40)
        worksheet4.freeze_panes(1, 1)  
    
    if not df.empty:
        df.to_excel(writer, index=False, sheet_name="Chat")
        worksheet = writer.sheets['Chat']       
        worksheet.set_column('A:A', 19)
        worksheet.set_column('B:B', 25)
        worksheet.set_column('C:C', 15)
        worksheet.set_column('D:D', 15)
        worksheet.set_column('E:E', 45)
        worksheet.set_column('F:F', 20)
        worksheet.set_column('G:G', 35)
        worksheet.set_column('H:H', 20)
        worksheet.freeze_panes(1, 1) 
    
    if not df2.empty:
        df2.to_excel(writer, index=False, sheet_name="Group Chat")
        worksheet2 = writer.sheets['Group Chat']       
        worksheet2.set_column('A:A', 19)
        worksheet2.set_column('B:B', 15)
        worksheet2.set_column('C:C', 15)
        worksheet2.set_column('D:D', 25)
        worksheet2.set_column('E:E', 40)
        worksheet2.set_column('F:F', 45)
        worksheet2.set_column('G:G', 30)
        worksheet2.set_column('H:H', 50)
        worksheet2.freeze_panes(1, 1)  

    if not df5.empty:
        df5.to_excel(writer, index=False, sheet_name="Friends List")
        worksheet5 = writer.sheets['Friends List']       
        worksheet5.set_column('A:A', 40)
        worksheet5.freeze_panes(1, 1)   

    if not df6.empty:
        df6.to_excel(writer, index=False, sheet_name="Snap")
        worksheet6 = writer.sheets['Snap']       
        worksheet6.set_column('A:A', 19)
        worksheet6.set_column('B:B', 19)
        worksheet6.set_column('C:C', 19)
        worksheet6.set_column('D:D', 23)
        worksheet6.set_column('E:E', 21)
        worksheet6.set_column('F:F', 20)
        worksheet6.set_column('G:G', 15)
        worksheet6.set_column('H:H', 15)
        worksheet6.set_column('I:I', 20)
        worksheet6.set_column('J:J', 20)
        worksheet6.set_column('K:K', 20)
        worksheet6.freeze_panes(1, 1) 
        
    if not df7.empty:
        df7.to_excel(writer, index=False, sheet_name="Group Story Snap")
        worksheet7 = writer.sheets['Group Story Snap']       
        worksheet7.set_column('A:A', 19)
        worksheet7.set_column('B:B', 35)
        worksheet7.set_column('C:C', 35)
        worksheet7.set_column('D:D', 28)
        worksheet7.set_column('E:E', 20)
        worksheet7.set_column('F:F', 20)
        worksheet7.freeze_panes(1, 1)  
        
    if not df8.empty:
        df8.to_excel(writer, index=False, sheet_name="Group")
        worksheet8 = writer.sheets["Group"]       
        worksheet8.set_column('A:A', 35)
        worksheet8.set_column('B:B', 35)
        worksheet8.set_column('C:C', 30)
        worksheet8.freeze_panes(1, 1)                 
                
    if not df9.empty:
        df9.to_excel(writer, index=False, sheet_name="Story")
        worksheet9 = writer.sheets['Story']       
        worksheet9.set_column('A:A', 19)
        worksheet9.set_column('B:B', 19)
        worksheet9.set_column('C:C', 19)
        worksheet9.set_column('D:D', 23)
        worksheet9.set_column('E:E', 21)
        worksheet9.set_column('F:F', 20)
        worksheet9.set_column('G:G', 15)
    writer.save()
    writer.close()

saveBox()
saveXLSX()
fileexist()
sys.exit()
