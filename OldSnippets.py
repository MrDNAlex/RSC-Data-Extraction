
#AtoL = 12 # 12 Columns
#TeamsPerRow = AtoL - 3



#TierInfo = []
#TierInfo.append({"matchSheetName": schedulePages[0], "scrimSheetName" : schedulePages[1], "skiprowsMatch": 5, "usecolsMatch": "B:I", "nrowsMatch": 49 - 6,       "skiprowsScrim": 13, "usecolsScrim": "B:L", "nrowsScrim": 49-14, "preseasonDays": 3, "matchDays": 16, "playoffDays": 2, "spacing": 4, "matchDaysScrim": 16, "spacingScrim" : 2, "hasLunarConference": True, "hasDivisions": False})
#TierInfo.append({"matchSheetName": schedulePages[0], "scrimSheetName" : schedulePages[2], "skiprowsMatch": 56, "usecolsMatch": "B:K", "nrowsMatch": 102 - 57,    "skiprowsScrim": 13, "usecolsScrim": "B:L", "nrowsScrim": 49-14,  "preseasonDays": 3,"matchDays": 17, "playoffDays": 2,"spacing": 4,  "matchDaysScrim": 16, "spacingScrim" : 2, "hasLunarConference": True, "hasDivisions": False})
#TierInfo.append({"matchSheetName": schedulePages[0], "scrimSheetName" : schedulePages[3], "skiprowsMatch": 109, "usecolsMatch": "B:Q", "nrowsMatch": 158 - 110,  "skiprowsScrim": 13, "usecolsScrim": "B:L", "nrowsScrim": 87-14, "preseasonDays": 3, "matchDays": 18, "playoffDays": 3, "spacing": 5, "matchDaysScrim": 16, "spacingScrim" : 2, "hasLunarConference": True, "hasDivisions": False})
#TierInfo.append({"matchSheetName": schedulePages[0], "scrimSheetName" : schedulePages[4], "skiprowsMatch": 167, "usecolsMatch": "B:V", "nrowsMatch": 218 - 168,  "skiprowsScrim": 13, "usecolsScrim": "B:L", "nrowsScrim": 87-14, "preseasonDays": 3, "matchDays": 19, "playoffDays": 3, "spacing": 5, "matchDaysScrim": 16, "spacingScrim" : 2, "hasLunarConference": True, "hasDivisions": True})
#TierInfo.append({"matchSheetName": schedulePages[0], "scrimSheetName" : schedulePages[5], "skiprowsMatch": 227, "usecolsMatch": "B:V", "nrowsMatch": 278 - 228,  "skiprowsScrim": 13, "usecolsScrim": "B:L", "nrowsScrim": 87-14, "preseasonDays": 3, "matchDays": 19, "playoffDays": 3, "spacing": 5, "matchDaysScrim": 16, "spacingScrim" : 2, "hasLunarConference": True, "hasDivisions": True})
#TierInfo.append({"matchSheetName": schedulePages[0], "scrimSheetName" : schedulePages[6], "skiprowsMatch": 285, "usecolsMatch": "B:Q", "nrowsMatch": 334 - 286,  "skiprowsScrim": 13, "usecolsScrim": "B:L", "nrowsScrim": 87-14, "preseasonDays": 3, "matchDays": 18, "playoffDays": 3, "spacing": 5, "matchDaysScrim": 16, "spacingScrim" : 2, "hasLunarConference": True, "hasDivisions": False})
#TierInfo.append({"matchSheetName": schedulePages[0], "scrimSheetName" : schedulePages[7], "skiprowsMatch": 341, "usecolsMatch": "B:K", "nrowsMatch": 387 - 342,  "skiprowsScrim": 13, "usecolsScrim": "B:L", "nrowsScrim": 49-14, "preseasonDays": 3, "matchDays": 17, "playoffDays": 2, "spacing": 4, "matchDaysScrim": 16, "spacingScrim" : 2, "hasLunarConference": True, "hasDivisions": False})
#TierInfo.append({"matchSheetName": schedulePages[0], "scrimSheetName" : schedulePages[8], "skiprowsMatch": 394, "usecolsMatch": "B:H", "nrowsMatch": 438 - 395,  "skiprowsScrim": 13, "usecolsScrim": "B:L", "nrowsScrim": 49-14, "preseasonDays": 3, "matchDays": 16, "playoffDays": 3, "spacing": 4, "matchDaysScrim": 16, "spacingScrim" : 2, "hasLunarConference": True, "hasDivisions": False})
#TierInfo.append({"matchSheetName": schedulePages[0], "scrimSheetName" : schedulePages[9], "skiprowsMatch": 445, "usecolsMatch": "B:K", "nrowsMatch": 464 - 446,  "skiprowsScrim": 13, "usecolsScrim": "B:k", "nrowsScrim": 30-14, "preseasonDays": 3, "matchDays": 15, "playoffDays": 1, "spacing": 4, "matchDaysScrim": 16, "spacingScrim" : 2, "hasLunarConference": False, "hasDivisions": False})


#for i in range(len(Tiers)):
#    MatchScheduleSheets[Tiers[i]] = LinearizeMatchDaySchedule(ScheduleSheet, TierInfo[i])
#    
#    divisionOffset = 3 if TierInfo[i]["hasDivisions"] else 0
#    conferenceMultiple = 2 if TierInfo[i]["hasLunarConference"] else 1
#    
#    teamColumnNums = len(GetTable(ScheduleSheet, TierInfo[i]["matchSheetName"], TierInfo[i]["skiprowsMatch"], TierInfo[i]["usecolsMatch"], TierInfo[i]["nrowsMatch"]).columns)
#    
#    TierInfo[i]["teamNum"] = (teamColumnNums - 2 - divisionOffset) * conferenceMultiple
#    TierInfo[i]["scrimRows"] = int(TierInfo[i]["teamNum"] / TeamsPerRow) + 1
#    
#    ScrimScheduleSheets[Tiers[i]] = LinearizeScrimDaySchedule(ScheduleSheet, TierInfo[i])
