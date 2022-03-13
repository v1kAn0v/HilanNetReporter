import PySimpleGUI


class SimpleForm:

    def __init__(self):
        self.event: object
        self.values: list

    @classmethod
    def submit(cls):
        PySimpleGUI.theme('DarkAmber')

        layout = [[PySimpleGUI.Text('הזן יום, לקוח, פרויקט, משימה, כניסה, יציאה, פירוט עבודה.')],
                  [PySimpleGUI.InputText()],
                  [PySimpleGUI.Submit(), PySimpleGUI.Cancel()]]

        window = PySimpleGUI.Window('Hilan Daily Reports', layout)

        cls.event, cls.values = window.read()
        print(cls.values)
        window.close()

    @classmethod
    def popup(cls):
        text_input = cls.values[0]
        PySimpleGUI.popup('You reported', text_input)
        return text_input
