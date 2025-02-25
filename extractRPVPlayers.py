import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import re
import rscapi
from rscapi.models.league_players_list200_response import LeaguePlayersList200Response
from rscapi.rest import ApiException
import os

# Ignore the first few rows of the file
RPVTable = pd.read_csv('RSCC S22 RPV - Sorted RPV.csv', skiprows=4 )

# Extract only the Rival Tier Players
RPVTable = RPVTable[RPVTable["Tier"] == "Rival"]

# Useful Data Extraction
RPVData = { "PlayerName" : 2, "RSCID" : 1, "Tier": 0, "Team" : 3, "Conference" : 4, "SBV" : 5 , "IDR" : 6, "RPV" : 7}

def sanitize_filename(filename, replacement="_"):
    # Define invalid characters for Windows file names
    invalid_chars = r'[\\/:*?"<>|]'
    # Replace them with the given replacement character (default: "_")
    return re.sub(invalid_chars, replacement, filename)

for i in range(0, len(RPVTable)):
    playerName = sanitize_filename(RPVTable.iloc[i][RPVData["PlayerName"]])
    obsidianFileName = playerName + ".md"
    
    obsidianFileContent = "---\n"
    
    for data in RPVData:
        obsidianFileContent += f"{data}: {RPVTable.iloc[i][RPVData[data]]}\n"
        
    obsidianFileContent += "---\n"
    
    with open(f"Output/{obsidianFileName}", "w") as file:
        file.write(obsidianFileContent)
    
