class Program:
    def __init__(self, school, program, psych):
        self._school = school
        self._program = program
        self._psych = psych

    def __str__(self):
        return "The " + self._school + " offers the " + self._program + " program. It is supported by " + self._psych + "."

    def get_school(self):
        """Return the school that offers the program."""
        return self._school

    def get_program(self):
        """Return the type of program that is offered by the school."""
        return self._program

    def get_psych(self):
        """Return the psychologist who supports the program."""
        return self._psych