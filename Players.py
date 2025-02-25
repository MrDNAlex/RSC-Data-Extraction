class Player:
    def __init__(self, ID, name, team, role, conference, RPV):
        self.ID = ID
        self.name = name
        self.team = team
        self.role = role
        self.conference = conference
        self.RPV = RPV

    def __str__(self):
        return f"{self.name} ({self.team}) - {self.role} - {self.RPV}"

    def PrintObsidian(self):
        
        fileContent = "---\n"
        
        # Fill in all the Info here
        
        fileContent +- "---\n"
        
        return fileContent