class Staff:
    """A class to represent a staff member."""
    def __init__(self, name, job, fte, team, sped_programs=[], beh_programs=[]):
        self._name = name
        self._job = job
        self._fte = fte
        self._team = team
        self._sped_programs = sped_programs
        self._beh_programs = beh_programs

    def __str__(self):
        return self._name + " is a " + self._job + " on the " + self._team + " team. They have an FTE of " + str(self._fte) + ", and their assigned programs are " + str(self._sped_programs) + " and " + str(self._beh_programs) + "."

    def get_name(self):
        """Return the name of the staff member."""
        return self._name

    def get_job(self):
        """Return the job of the staff member."""
        return self._job

    def get_fte(self):
        """Return the FTE of the staff member."""
        return self._fte

    def get_team(self):
        """Return the team of the staff member."""
        return self._team

    def get_programs(self):
        """Return the programs the staff member is assigned to."""
        return self._sped_programs, self._beh_programs

    def set_program(self, program):
        """Add a program to a staff member's list of programs."""
        if program.team == "SPED":
            self._sped_programs.append(program)
        elif program.team == "Behaviour":
            self._beh_programs.append(program)