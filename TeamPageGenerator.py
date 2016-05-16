#Code by Luke Farritor
#Project started on Mar 1, 2016

#Imports
from openpyxl.cell import get_column_letter, column_index_from_string
from string import Template
import openpyxl
import codecs
import os

#Variables
#Team Page Template Perams
templateName = "Template.html"
teamFileName = "Teams.html"

#Parsed Score Spreadsheet Perams
parsedScoresName = "Parsed.xlsx"
rowsInScores = 4765

#Team Directory Perams
teamDirLine = '<p><a href="$number.html">$number</a> - Avg: $avg \n' #the line per team
teamDirTemplate = Template(teamDirLine) #make it into a template
teamDirectory = ""
teamDirHtmlStartFl = "TeamDirStart.html"
teamDirHtmlEndFl = "TeamDirEnd.html"
teamDirHtmlFl = "TeamDirectory.html"

#Loading Scoring Sheet
print("Loading Workbook... (This will take a moment)")
scoresFile = openpyxl.load_workbook(parsedScoresName)
print("Loading Sheet...")
scrsSht = scoresFile.get_sheet_by_name('Sheet')

#Load Team Page HTML Template
print("Loading HTML Template...")
templateHtml = codecs.open(templateName, 'r')
template = Template(templateHtml.read())
templateHtml.close()


#Load Global Stats Page
print("Loading Stats Template...")
statPageTemplateName = "indexTemplate.html"
statPageName = "index.html"
statTemplateHtml = codecs.open(statPageTemplateName, 'r')
statTemplate = Template(statTemplateHtml.read())
statTemplateHtml.close()

#Load Team Directory HTML Template
print("Loading Team Page HTML...")
teamDirStartFl = codecs.open(teamDirHtmlStartFl, 'r')
teamDirectory = teamDirStartFl.read()
teamDirStartFl.close()

teamDirEndFl = codecs.open(teamDirHtmlEndFl, 'r')
teamDirEnd = teamDirEndFl.read()
teamDirEndFl.close()

#Variables for calculating avg win/loss score
totalWins = 0
totalWinAcc = 0
avgWinScore = 0

global teamAvg #stored globally
teamAvg = 0

####################################################################################################

def percentage(part, whole):
        if(part * whole != 0): #if one of them is not zero
                return 100 * float(part)/float(whole)
        else:
                return 0

####################################################################################################

def avgWinScore(rows):
        totalWinAcc = 0
        for n in range(2, rows):
                totalWinAcc += int(str(scrsSht['O' + str(n)].value))
        return totalWinAcc / (rows - 1)

####################################################################################################

def avgScore(rows):
        totalScrAcc = 0
        for n in range(2, rows):
                totalScrAcc += int(str(scrsSht['H' + str(n)].value))
                totalScrAcc += int(str(scrsSht['L' + str(n)].value))
        return totalScrAcc / (rows * 2 - 1)

####################################################################################################

def getTeamList(rows):
        teams = ','
        for n in range(2, rows):
                if(teams.find(str(scrsSht['E' + str(n)].value)) == -1): #if the team is not in the team directory...
                        teams += str(scrsSht['E' + str(n)].value) + ','
                
                if(teams.find(str(scrsSht['F' + str(n)].value)) == -1): #if the team is not in the team directory...
                        teams += str(scrsSht['F' + str(n)].value) + ','

                if(teams.find(str(scrsSht['I' + str(n)].value)) == -1): #if the team is not in the team directory...
                        teams += str(scrsSht['I' + str(n)].value) + ','

                if(teams.find(str(scrsSht['J' + str(n)].value)) == -1): #if the team is not in the team directory...
                        teams += str(scrsSht['J' + str(n)].value) + ','
        return teams

####################################################################################################

def getTeamInfo(number, tWins, tWinsAcc, teamList):
        exists = False #weather the team exists
        teamAppearances = 0 #number of matches
        teamAlliance = 'r' #what allance they were on
        teamWins = 0
        teamLosses = 0
        teamTies = 0
        teamWinPercentage = 0
        teamTotalScores = 0
        teamAvgScore = 0
        teamHighest = 0
        teamScore = 0
        teamScores = [0] * 10000000
        winner = ''
        
        if(teamList.find(str(number) + ',') == -1):
                return False
        
        for row in range(1, rowsInScores): #increment through the rows of data
                if(str(scrsSht['N' + str(row)].value).find(str(number)+',') >= 0): #checks if the team competed in this row/match
                        exists = True #if we can find the team, the team exists
                        teamAppearances += 1 #add one to the amount of team matches
                        winner = str(scrsSht['M' + str(row)].value)
                        matchSummCell = scrsSht['N' + str(row)].value

                        #print(str(matchSummCell.find(str(number)))+ ', ' + str(matchSummCell.find('vs')))
                        
                        if(int(matchSummCell.find(str(number))) >= int(matchSummCell.find('vs'))): #finds if they were on red or blue
                                teamAlliance = 'b'
                        else:
                                teamAlliance = 'r'
                        if(teamAlliance == winner): #if team won
                                teamWins += 1
                        elif(teamAlliance != 't'): #the team lost
                                teamLosses += 1
                        else:
                                teamTies += 1

                        if(teamAlliance == 'r'): #if we are on red..
                                teamTotalScores += scrsSht['H' + str(row)].value #...add the red score to the total scores
                                teamScore = scrsSht['H' + str(row)].value
                        else:
                                teamTotalScores += scrsSht['L' + str(row)].value #add blue score to total scores
                                teamScore = scrsSht['L' + str(row)].value

                        if(teamHighest < teamScore):
                                teamHighest = teamScore
                        
                        if(winner == 'r'):
                                tWinsAcc += scrsSht['H' + str(row)].value #add the red score to the global winning scores
                                tWins += 1
                        elif(winner == 'b'):
                                tWinsAcc += scrsSht['L' + str(row)].value #add blue score to the global winning scores
                                tWins += 1
                                


        if(exists):
                teamAvgScore = int(teamTotalScores / teamAppearances)
                teamWinPercentage = int(percentage(teamWins, teamAppearances))
        else:
                print(str(number) + ' does not exist')
                return False

        htmlSubs = dict(
                teamNumber=number,
                avgScore=teamAvgScore,
                matches=teamAppearances,
                wins=teamWins,
                ties=teamTies,
                losses=teamLosses,
                matchWin=teamWinPercentage,
                highest=teamHighest) #substitutions into team page
        
        teamPageStr = template.safe_substitute(htmlSubs)
        teamPage = codecs.open(str(number) + '.html', 'w+') #opens(or creates) team page
        if(teamPage.read() != ''):
                teamPage.close()
                os.remove(str(number) + '.html')
                teamPage = codecs.open(str(number) + '.html', 'w+')

        teamPage.write(teamPageStr)
        teamPage.close()
        
        
        print(str(number) + ',' + str(teamTies))
        return str(str(teamAvgScore) + ',tWins:' + str(tWins) + ',tWinAcc:' + str(tWinsAcc))

####################################################################################################

statPageSubs = dict(
        avgScore = int(avgScore(rowsInScores)),
        avgWinScore = int(avgWinScore(rowsInScores)),
        worldHigh = -1)

statPageStr = statTemplate.safe_substitute(statPageSubs)
statPage = codecs.open(statPageName, 'w+') #opens(or creates) stat page
if(statPage.read() != ''):
        statPage.close()
        os.remove(statPageName)
        statPage = codecs.open(statPageName, 'w+')

statPage.write(statPageStr)
statPage.close()

####################################################################################################

teamList = getTeamList(rowsInScores)
for n in range(4104,4108):
        teamInfo = getTeamInfo(int(n), totalWins, totalWinAcc, teamList) 
        if(teamInfo):
                teamDirSubs = dict(number=n, #team number
                        avg=teamInfo[:teamInfo.find(',')]) #find avg score 
                teamDirectory += teamDirTemplate.safe_substitute(teamDirSubs)

                totalWins = int(teamInfo[teamInfo.find('tWins:') + 6:teamInfo.find(',', teamInfo.find('tWins:'))]) 
                totalWinAcc += int(teamInfo[teamInfo.find('tWinAcc:') + 8 : teamInfo.find(',', teamInfo.find('tWinAcc:'))])


teamDirectory += teamDirEnd

####################################################################################################

#write to team Dir file
teamDirPage = codecs.open(teamDirHtmlFl, 'w+') #opens(or creates) team Directory page
if(teamDirPage.read() != ''):
        teamDirPage.close()
        os.remove(teamDirHtmlFl)
        teamPage = codecs.open(teamDirHtmlFl, 'w+')

teamDirPage.write(teamDirectory)
teamDirPage.close()
