from gtts import gTTS

def text_to_speech(myText, output_file="static/output.mp3"):
    tts = gTTS(text=myText, lang='en', slow=False)
    tts.save(output_file)
    return output_file

