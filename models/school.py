class School:
    def __init__(self, name, area, school_psych, address, latitude_radian, longitude_radian, programs):
        self._name = name
        self._area = area
        self._school_psych = school_psych
        self._address = address
        self._latitude_radian = latitude_radian
        self._longitude_radian = longitude_radian
        self._programs = programs

    def __str__(self):
        return self._name + " is located in " + self._area + ". The school psychologist is " + self._school_psych + ". The address is " + self._address + ". The latitude is " + str(self._latitude_radian) + " and the longitude is " + str(self._longitude_radian) + ". The programs offered are " + str(self._programs) + "."

    def get_name(self):
        """Return the name of the school."""
        return self._name

    def get_area(self):
        """Return the city area of the school."""
        return self._area

    def get_school_psych(self):
        """Return the school psychologist who is assigned to that school."""
        return self._school_psych

    def get_address(self):
        """Return the address of the school."""
        return self._address

    def get_latitude_radian(self):
        """Return the latitude of the school in radians."""
        return self._latitude_radian

    def get_longitude_radian(self):
        """Return the longitude of the school in radians."""
        return self._longitude_radian

    def get_programs(self):
        """Return the Diverse Learning programs offered by the school."""
        return self._programs