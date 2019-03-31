# -*- coding: utf-8 -*-

import json
import re
import ast
import xlsxwriter
from datetime import datetime
from collections import Counter

date = datetime.now().strftime('%m_%d_%Y')
print (date)
day_of_year = datetime.now().timetuple().tm_yday
print('Day', day_of_year, 'of this year')

## OPEN JSON - GET RID OF EMOJIS - INITIATE VARS ##
f = open('theofficialmotiveapp-export.json', 'r+', encoding="utf-8")
motive_data = json.load(f)
motive_data_clean = {str(k) : re.sub(r'[^\x00-\x7F]',' ', str(v)) for k, v in motive_data.items()}
challenges = ast.literal_eval(motive_data_clean['challenges'])
users = ast.literal_eval(motive_data_clean['users'])
notifications = ast.literal_eval(motive_data_clean['notifications'])
friends = ast.literal_eval(motive_data_clean['friends'])
chat = ast.literal_eval(motive_data_clean['chat'])
activity = ast.literal_eval(motive_data_clean['activity'])
###################################################


## CREATE WORKBOOK & SHEETS ##
workbook = xlsxwriter.Workbook(date+'.xlsx')
S_Totals = workbook.add_worksheet('TOTALS')
S_Users = workbook.add_worksheet('USERS')
S_Challenges = workbook.add_worksheet('CHALLENGES')
S_Notifications = workbook.add_worksheet('NOTIFICATIONS')
S_Chat = workbook.add_worksheet('CHAT')
S_Feed = workbook.add_worksheet('FEED')
#############################

## CREATE FONT FORMATS ##
Bold14 = workbook.add_format({'bold':True, 'font_size': 14})
Bold14Green = workbook.add_format({'bold':True, 'font_size': 14, 'font_color': 'green'})
#########################

####       ####
#### USERS ####
####       ####
ActiveUsers = []
ActiveUserCount = 0
total_users = 0
total_comments = 0
total_likes = 0

for challenge_id in challenges.keys():
    for challenge_params in challenges[challenge_id].keys():
        if challenge_params == 'points':
            for points_entered_by in challenges[challenge_id]['points'].keys():
                for days_points_entered in challenges[challenge_id]['points'][points_entered_by].keys():
                    if datetime.now().toordinal() - datetime.strptime(days_points_entered,'%Y-%m-%d').toordinal() <= 14:
                        ActiveUsers.append(points_entered_by)

for challenge_id in chat.keys():
    for chat_id in chat[challenge_id].keys():
        if datetime.now().toordinal() - datetime.strptime(chat[challenge_id][chat_id]['date'][:10], '%Y-%m-%d').toordinal() <= 14:
            ActiveUsers.append(chat[challenge_id][chat_id]['uid'])

for user_id in notifications.keys():
    for nf_id in notifications[user_id].keys():
        if datetime.now().toordinal() - datetime.strptime(notifications[user_id][nf_id]['date'], '%Y-%m-%d').toordinal() <= 14:
            if notifications[user_id][nf_id]['read'] == 1:
                ActiveUsers.append(user_id)

for user_id in activity.keys():
    for activity_id in activity[user_id].keys():
        for properties in activity[user_id][activity_id].keys():
            if properties == 'comments':
                for comment_id in activity[user_id][activity_id][properties].keys():
                    total_comments += 1
                    for param in activity[user_id][activity_id][properties][comment_id].keys():
                        if datetime.now().toordinal() - datetime.strptime(activity[user_id][activity_id][properties][comment_id]['time'][:10], '%Y-%m-%d').toordinal() <= 14:
                            ActiveUsers.append(activity[user_id][activity_id][properties][comment_id]['uid'])
            elif properties == 'likes':
                for uid in activity[user_id][activity_id][properties].keys():
                    total_likes += 1
            
ActiveUsersFinal = list(set(ActiveUsers))
#print(ActiveUsersFinal)
print(len(ActiveUsersFinal), 'users have been active in the last two weeks')

S_Users.set_column('A:C', 15)
S_Users.set_column('E:G', 15)

for key in users.keys():
    total_users = total_users + 1

S_Users.write('A1', 'Active Users', Bold14)
S_Users.write('B1', len(ActiveUsersFinal), Bold14Green)
S_Users.write('A2', 'Total Users', Bold14)
S_Users.write('B2', total_users, Bold14Green)

S_Users.write('D1', 'User', Bold14)
x = 2
for key in users.keys():
    S_Users.write('D%d'%x, users[key]['firstname'])
    S_Users.write('E%d'%x, users[key]['lastname'])
    S_Users.write('F%d'%x, users[key]['email'])
    x = x + 1

S_Users.write('H1', 'Active Users', Bold14)
x = 2
for k in ActiveUsersFinal:
    if k in users.keys():
            S_Users.write('H%d'%x, users[k]['firstname'])
            S_Users.write('I%d'%x, users[k]['lastname'])
            S_Users.write('J%d'%x, users[k]['email'])
            x = x + 1

D_user_id = {}
def CountFriends(user_id):
    D_user_id[user_id] = len(friends[user_id].keys())

for user_id in friends.keys():
    CountFriends(user_id)

L_friend_count = []
for user_id in D_user_id.keys():
    L_friend_count.append(D_user_id[user_id])

average_friends_per_user = sum(L_friend_count)/len(L_friend_count)
average_friends_per_user = round(average_friends_per_user,3)
##print(average_friends_per_user)
S_Users.write('A5', 'Avg Friends/User', Bold14)
S_Users.write('B5', average_friends_per_user, Bold14Green)
              
L_activefriends_count = []
try: 
    for user_id in ActiveUsersFinal:
        L_activefriends_count.append(D_user_id[user_id])
except:
    pass

average_friends_per_activeuser = sum(L_activefriends_count)/len(L_activefriends_count)
average_friends_per_activeuser = round(average_friends_per_activeuser,3)
##print(average_friends_per_activeuser)
S_Users.write('A6', 'Avg Friends/ActiveUser', Bold14)
S_Users.write('B6', average_friends_per_activeuser, Bold14Green)

#######################################################
#######################################################
#######################################################



####            ####
#### CHALLENGES ####
####            ####
S_Challenges.set_column('A:G', 15)


S_Challenges.write('D1', 'Challenge Name', Bold14)
S_Challenges.write('E1', 'Type', Bold14)
S_Challenges.write('F1', 'User Count', Bold14)

ActiveChallenges = []
active_challenges = 0
total_challenges = 0
checklist_challenges = 0
x = 2
y = 2
for key in challenges.keys():
        encoded = challenges[key]['motiveName'].encode('utf-8').strip()
##        S_Challenges.write('D%d'%x, challenges[key]['motiveName'])
        total_challenges = total_challenges + 1
        x = x + 1
        LastUpdated = challenges[key].get('lastUpdated')
        if LastUpdated != None:
             if datetime.now().toordinal() - datetime.strptime(LastUpdated,'%Y-%m-%d').toordinal() <= 14:
                    S_Challenges.write('D%d'%y, challenges[key]['motiveName'])
                    S_Challenges.write('E%d'%y, challenges[key]['type'])
                    S_Challenges.write('F%d'%y, len(challenges[key]['challengeUsers']))
                    ActiveChallenges.append(len(challenges[key]['challengeUsers']))
                    if challenges[key]['type'] == 'Checklist':
                        checklist_challenges = checklist_challenges + 1
                    y = y + 1
                    active_challenges = active_challenges + 1

UsersPerActiveChallenge = workbook.add_chart({'type': 'column'})


ActiveChallengesHistData = []
for key, value in Counter(ActiveChallenges).items():
    ActiveChallengesHistData.append(int(value))

ActiveChallengesHistData = sorted(ActiveChallengesHistData)
S_Challenges.write_column('Z1', ActiveChallengesHistData)
UsersPerActiveChallenge.add_series({'values': '=CHALLENGES!$Z$1:$Z$'+str(len(ActiveChallengesHistData))})
UsersPerActiveChallenge.set_title({'name': 'Number of Users per Active Challenge'})
UsersPerActiveChallenge.set_legend({'none': True})
UsersPerActiveChallenge.set_x_axis({'name': 'User Count'})
UsersPerActiveChallenge.set_y_axis({'name': '# of Active Challenges'})
S_Challenges.insert_chart('I1', UsersPerActiveChallenge)


S_Challenges.write('A1', 'Active Graphs', Bold14)
S_Challenges.write('B1', active_challenges-checklist_challenges, Bold14Green)

S_Challenges.write('A2', 'Active Checklists', Bold14)
S_Challenges.write('B2', checklist_challenges, Bold14Green)

S_Challenges.write('A3', 'Total Active', Bold14)
S_Challenges.write('B3', active_challenges, Bold14Green)

S_Challenges.write('A4', 'Overall Total', Bold14)
S_Challenges.write('B4', total_challenges, Bold14Green)

total_points = 0
dates_points_entered = []
for challenge_id in challenges.keys():
    if 'lastUpdated' in challenges[challenge_id].keys():
        if 'points' in challenges[challenge_id].keys():
            for user_id in challenges[challenge_id]['points'].keys():
                for date_entered in challenges[challenge_id]['points'][user_id].keys():
                    total_points = total_points + 1
                    dates_points_entered.append(date_entered)

total_steps = 0
step_completions = 0
for challenge_id in challenges.keys():
    if 'lastUpdated' in challenges[challenge_id].keys():
            if 'steps' in challenges[challenge_id].keys():
                for step_id in challenges[challenge_id]['steps'].keys():
                    total_steps = total_steps + 1
                    if 'completed' in challenges[challenge_id]['steps'][step_id].keys():
                        for user_id in challenges[challenge_id]['steps'][step_id]['completed'].keys():
##                            print (user_id)
                            step_completions = step_completions + 1
    

print (total_points, 'total points have been entered so far')
print (total_steps, 'steps have been created so far')
print (step_completions, 'step completions have been made so far')

S_Challenges.write('A7', 'Graph Points', Bold14)
S_Challenges.write('B7', total_points, Bold14Green)
S_Challenges.write('A8', 'Steps Created', Bold14)
S_Challenges.write('B8', total_steps, Bold14Green)
S_Challenges.write('A9', 'Step Completions', Bold14)
S_Challenges.write('B9', step_completions, Bold14Green)

dates_points_entered = sorted(dict(Counter(dates_points_entered)).items())
dates_points_entered = dates_points_entered[137:]

point_date_keys = []
point_date_values = []
for point_date in dates_points_entered:
    point_date_keys.append(point_date[0])
    point_date_values.append(point_date[1])

PointsPerDay = workbook.add_chart({'type': 'line'})

S_Challenges.write_column('AB1', point_date_keys)
S_Challenges.write_column('AC1', point_date_values)
PointsPerDay.add_series({'values': '=CHALLENGES!$AB$1:$AB$'+str(len(point_date_keys))})
PointsPerDay.add_series({'values': '=CHALLENGES!$AC$1:$AC$'+str(len(point_date_values))})
PointsPerDay.add_series({
    'values':    '=CHALLENGES!$AC$1:$AC$'+str(len(point_date_values)),
    'trendline': {'type': 'polynomial', 'order': 2},
})

PointsPerDay.set_title({'name': 'Points Per Day'})
##PointsPerDay.set_legend({'none': True})
PointsPerDay.set_x_axis({'date_axis': True, 'name': 'Day'})
PointsPerDay.set_y_axis({'name': '# of Points'})
PointsPerDay.set_legend({'delete_series': [0, 1]})
S_Challenges.insert_chart('I17', PointsPerDay)

#######################################################
#######################################################
#######################################################


####               ####
#### NOTIFICATIONS ####
####               ####

total_nf = 0
total_fr = 0
total_cr = 0
total_cu = 0
total_cm = 0
total_cj = 0
total_ia = 0
total_fra = 0
total_pa = 0
total_dc = 0
total_ce = 0
total_du = 0
total_de = 0

text_cj = "has joined the challenge"
text_image = "has added an image"
text_progress = "has added progress"
text_datechange = "date has been changed"
text_expire = "you can still extend it"
text_updated = "has been updated to"
text_extended = "been extended until"

NF_dates = []

for user_id in notifications.keys():
    for nf_id in notifications[user_id].keys():
        total_nf = total_nf + 1
        NF_dates.append(notifications[user_id][nf_id]['date'])
        if notifications[user_id][nf_id]['type'] == 'friendRequest':
            total_fr = total_fr + 1
        elif notifications[user_id][nf_id]['type'] == 'challengeRequest':
            total_cr = total_cr + 1
        elif notifications[user_id][nf_id]['type'] == 'challengeUpdate':
            total_cu = total_cu + 1
        elif notifications[user_id][nf_id]['type'] == 'challengeMessage':
            total_cm = total_cm + 1
        elif notifications[user_id][nf_id]['type'] == 'friendRequestAccepted':
            total_fra = total_fra + 1
        if text_cj in notifications[user_id][nf_id]['text']:
            total_cj = total_cj + 1
        if text_image in notifications[user_id][nf_id]['text']:
            total_ia = total_ia + 1
        if text_progress in notifications[user_id][nf_id]['text']:
            total_pa = total_pa + 1
        if text_datechange in notifications[user_id][nf_id]['text']:
            total_dc = total_dc + 1
        if text_expire in notifications[user_id][nf_id]['text']:
            total_ce = total_ce + 1
        if text_updated in notifications[user_id][nf_id]['text']:
            total_du = total_du + 1



NF_dates = dict(Counter(NF_dates))
NF_dates = sorted(NF_dates.items())
#print(NF_dates[101])
NF_dates = NF_dates[101:]
#print(NF_dates)

NF_dates_values = []
NF_dates_keys = []
for date in NF_dates:
    NF_dates_keys.append(date[0])
    NF_dates_values.append(date[1])

#print(NF_dates_keys)
#print(NF_dates_values)


NFsPerDay = workbook.add_chart({'type': 'line'})

S_Notifications.write_column('Y1', NF_dates_keys)
S_Notifications.write_column('Z1', NF_dates_values)
NFsPerDay.add_series({'values': '=NOTIFICATIONS!$Y$1:$Y$'+str(len(NF_dates_keys))})
NFsPerDay.add_series({'values': '=NOTIFICATIONS!$Z$1:$Z$'+str(len(NF_dates_values))})
NFsPerDay.add_series({
    'values':    '=NOTIFICATIONS!$Z$1:$Z$'+str(len(NF_dates_values)),
    'trendline': {'type': 'exponential'},
})
NFsPerDay.set_title({'name': 'Notifications Per Day'})
NFsPerDay.set_legend({'none': True})
NFsPerDay.set_x_axis({'date_axis': True, 'name': 'Day'})
NFsPerDay.set_y_axis({'name': '# of Notifications'})
S_Notifications.insert_chart('D1', NFsPerDay)

print (total_nf, 'total notifications have been sent so far')
print (total_fr, 'friend requests nfs have been sent so far')
print (total_fra, 'friend request accepted nfs have been sent so far')
print (total_cr, 'challenge requests nfs have been sent so far')
print (total_cm, 'challenge message nfs have been sent so far')
print (total_cj, 'challenge joined nfs have been sent so far')
print (total_ia, 'image added nfs have been sent so far')
print (total_pa, 'progress added nfs have been sent so far')
print (total_dc, 'date change nfs have been sent so far')
print (total_ce, 'challenge expired nfs have been sent so far')
print (total_du, 'challenge date update nfs have been sent so far')
    
print (total_cu, 'challenge update nfs have been sent so far')

S_Notifications.set_column('A:B', 20)

S_Notifications.write('A1', 'Total', Bold14)
S_Notifications.write('B1', total_nf, Bold14Green)

S_Notifications.write('A2', 'Progress NFs', Bold14)
S_Notifications.write('B2', total_pa, Bold14Green)

S_Notifications.write('A3', 'Message NFs', Bold14)
S_Notifications.write('B3', total_cm, Bold14Green)

S_Notifications.write('A4', 'Friend Request NFs', Bold14)
S_Notifications.write('B4', total_fr, Bold14Green)

S_Users.write('A7', 'Friend Requests Accepted', Bold14)
S_Users.write('B7', total_fra, Bold14Green)

S_Notifications.write('A5', 'Image NFs', Bold14)
S_Notifications.write('B5', total_ia, Bold14Green)

#######################################################
#######################################################
#######################################################


####      ####
#### CHAT ####
####      ####

messages_sent = 0
date_sent = []

for challenge_id in chat.keys():
    for message_id in chat[challenge_id].keys():
        messages_sent = messages_sent + 1
        date_sent.append(chat[challenge_id][message_id]['date'][:10])
        
date_sent = sorted(dict(Counter(date_sent)).items())
date_sent_keys = []
date_sent_values = []
for date in date_sent:
    date_sent_keys.append(date[0])
    date_sent_values.append(date[1])

##print(date_sent_keys)
##print(date_sent_values)

S_Chat.set_column('A:B', 20)

MessagesPerDay = workbook.add_chart({'type': 'line'})

S_Chat.write_row('Y1', ['Date', 'Value'], Bold14)
S_Chat.write_column('Y2', date_sent_keys)
S_Chat.write_column('Z2', date_sent_values)
##print(str(len(date_sent_values)))
MessagesPerDay.add_series({'values': '=CHAT!$Y$2:$Y$'+str(len(date_sent_keys)+1)})
MessagesPerDay.add_series({'values': '=CHAT!$Z$2:$Z$'+str(len(date_sent_values)+1)})
MessagesPerDay.set_title({'name': 'Messages Per Day'})
MessagesPerDay.set_legend({'none': True})
MessagesPerDay.set_x_axis({'date_axis': True, 'name': 'Day (beginning 11/12/18)'})
MessagesPerDay.set_y_axis({'name': '# of Messages Sent'})
S_Chat.insert_chart('D1', MessagesPerDay)
S_Chat.write('A1', 'Total Messages Sent', Bold14)
S_Chat.write('B1', messages_sent, Bold14Green)

#######################################################
#######################################################
#######################################################


####        ####
#### TOTALS ####
####        ####
row = 0
col = 0
rowHeaders = ['Day',
              'Total Nfs',
              'Trailing ΔNF/Δt',
              'Trailing ΔNF/Δt percent change',
              'User total',
              'Active Users',
              'Aggregate Retention Rate',
              'Aggregate Retention Rate percent change',
              'Friend Connections',
              'Total Challenges',
              'Active Challenges',
              'Points',
              'Messages',
              'Images',
              'Active Graph Challenges',
              'Active Checklist Challenges',
              'Steps Created',
              'Steps Completed',
              'NF-Progress added',
              'NF-Challenge joined',
              'NF-Challenge request',
              'NF-Challenge message',
              'NF-Image added',
              'NF-Date update',
              'NF-challenge expired',
              'NF-Friend request',
              'Likes',
              'Comments']
rowValues = [day_of_year,
             total_nf,
             None,
             None,
             total_users,
             len(ActiveUsersFinal),
             None,
             None,
             total_fra,
             total_challenges,
             active_challenges,
             total_points,
             messages_sent,
             total_ia,
             active_challenges-checklist_challenges,
             checklist_challenges,
             total_steps,
             step_completions,
             total_pa,
             total_cj,
             total_cr,
             total_cm,
             total_ia,
             total_du,
             total_ce,
             total_fr,
             total_likes,
             total_comments]
S_Totals.write_row(row, col, tuple(rowHeaders))
row += 1
S_Totals.write_row(row, col, tuple(rowValues))


#######################################################
#######################################################
#######################################################


####                       ####
#### ACTIVITY INTERACTIONS ####
####                       ####
print (total_comments, 'comments have been made so far')
print (total_likes, 'likes have been made so far')




workbook.close()
