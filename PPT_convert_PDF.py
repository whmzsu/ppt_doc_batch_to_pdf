import comtypes.client
import os


def init_powerpoint():
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    return powerpoint


def ppt_to_pdf(powerpoint, inputFileName, outputFileName, formatType=32):
    deck = powerpoint.Presentations.Open(inputFileName)
    deck.SaveAs(outputFileName, formatType)  # formatType = 32 for ppt to pdf
    deck.Close()


def convert_files_in_folder(powerpoint, folder):
    for root, dirs, files in os.walk(folder):
        #files = os.listdir(folder)
        pptfiles = [f for f in files if f.endswith((".ppt", ".pptx"))]
        for pptfile in pptfiles:
            #fullpath = os.path.join(folder, pptfile)
            fullpath = os.path.join(root, pptfile)
            pdffullpath = os.path.join(
                root, os.path.splitext(pptfile)[0]+'.pdf')
            ppt_to_pdf(powerpoint, fullpath, pdffullpath)


if __name__ == "__main__":
    powerpoint = init_powerpoint()
    cwd = os.getcwd()
    convert_files_in_folder(powerpoint, cwd)
    powerpoint.Quit()
