from kivy.app import App
from kivy.uix.button import Button
from kivy.uix.floatlayout import FloatLayout
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput


class Interface(FloatLayout):
    def __init__(self,**kwargs):
        super().__init__(**kwargs)
        self.btn=Button(text= "CLICK ME", size_hint=(0.5,0.1), pos_hint={"center_x":0.5,"center_y":0.5})
        self.label=Label(text= "Hello", size_hint=(0.5,0.1), pos_hint={"center_x":0.5,"center_y":0.6})
        self.textInput=TextInput( size_hint=(0.5,0.1), pos_hint={"center_x":0.5,"center_y":0.7} ,multiline=False )
        self.textInput.bind(on_text_validate=self.button_press)

        self.btn.bind(on_press=self.button_press)
        self.add_widget(self.btn)
        self.add_widget(self.label)
        self.add_widget(self.textInput)




    def button_press(self, obj):
        print("button clicked....")
        data=self.textInput.text
        self.label.text=data






class ProjectApp(App):

    def build(self):
        return Interface()


ProjectApp().run()