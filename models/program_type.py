class ProgramType:
    def __init__(self, name, team, adaptive_func, cognitive_func, soc_emo_beh_func, phys_med_need, weight):
        self._team = team
        self._name = name
        self._adapt = adaptive_func
        self._cog = cognitive_func
        self._seb = soc_emo_beh_func
        self._phys_med = phys_med_need
        self._weight = weight

    def __str__(self):
        return "The " + str(self._name) + " program is managed by the " + str(self._team) + " team. It supports students with " + str(self._adapt) + " adaptive functioning deficits, " + str(self._cog) + " cognitive functioning deficits, and " + str(self._seb) + " social emotional behavior functioning deficts. It has an FTE weight of " + str(self._weight) + "."

    def get_name(self):
        """Return the name of the program."""
        return self._name

    def get_team(self):
        """Return the team of the program."""
        return self._team

    def get_adaptive_func(self):
        """Return the adaptive functioning level that the program supports."""
        return self._adapt

    def get_cog(self):
        """Return the cognitive functioning level that the program supports."""
        return self._cog

    def get_seb(self):
        """Return the social emotional behavioural functioning level that the program supports."""
        return self._seb

    def get_phys_med(self):
        """Return the physical medical needs that the program supports."""
        return self._phys_med

    def get_weight(self):
        """Return the FTE weight of the program."""
        return self._weight