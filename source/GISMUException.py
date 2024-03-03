class GISMUException(Exception):
    """Класс для ошибок ГИСМУ
        args[0] - пользовательское описание
        args[1] - ошибка"""

    def __init__(self, *args):
        if len(args) == 2:
            self.message = getattr(args[1], 'message', repr(args[1])) + "\n" + str(args[0])
        elif len(args) == 1:
            self.message = str(args[0])
        else:
            self.message = "Неизвестная ошибка"

    def __str__(self):
        return self.message

    def print_to_file(self, file_path):
        with open(file_path, 'a') as f:
            f.write(self.message + "\n\n")