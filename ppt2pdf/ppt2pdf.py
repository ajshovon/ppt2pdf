import win32com.client
import os
import argparse

parser = argparse.ArgumentParser(description='Convet ppt/pptx to pdf.')
parser.add_argument('-c', '--clean', help="Cleanup slide after converting to pdf.", action='store_true')
parser.add_argument('source', help="PPT/PPTX files", nargs='+')
args = parser.parse_args()
delete_ppt = False
pptFiles = args.source

if (args.clean):
    delete_ppt = True

def init_powerpoint():
    powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
    return powerpoint

def ppt_to_pdf(powerpoint, inputFileName, outputFileName, formatType = 32):
    if outputFileName[-3:] != 'pdf':
        outputFileName = os.path.splitext(outputFileName)[0] + ".pdf"
    deck = powerpoint.Presentations.Open(inputFileName)
    deck.SaveAs(outputFileName, formatType) # formatType = 32 for ppt to pdf
    deck.Close()
    if (delete_ppt):
        try:
            os.remove(inputFileName)
        except:
            print("Could not delete slide!")


if __name__ == "__main__":
    powerpoint = init_powerpoint()
    for ppt in pptFiles:
        ppt_to_pdf(powerpoint, ppt, ppt)
    powerpoint.Quit()
