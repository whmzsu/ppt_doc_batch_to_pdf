import comtypes.client
import os


def init_word():
    word = comtypes.client.CreateObject("Word.Application")
    word.Visible = 1
    return word


def ppt_to_pdf(word, inputFileName, outputFileName, formatType=17):
    deck = word.Documents.Open(inputFileName)
    deck.SaveAs(outputFileName, formatType)  # formatType = 17 for doc to pdf
    deck.Close()


def convert_files_in_folder(word, folder):
    for root, dirs, files in os.walk(folder):
        #files = os.listdir(folder)
        wordfiles = [f for f in files if f.endswith((".doc", ".docx"))]
        for wordfile in wordfiles:
            #fullpath = os.path.join(folder, pptfile)
            fullpath = os.path.join(root, wordfile)
            pdffullpath = os.path.join(
                root, os.path.splitext(wordfile)[0]+'.pdf')
            ppt_to_pdf(word, fullpath, pdffullpath)


if __name__ == "__main__":
    word = init_word()
    cwd = os.getcwd()
    convert_files_in_folder(word, cwd)
    word.Quit()
