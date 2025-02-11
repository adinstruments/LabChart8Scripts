
import win32com.client
import time

# ------------------------------------------------------------------------------

def invoke_com_method(com_object, method_name, *args):
    try:
        # Get the method dynamically
        method = getattr(com_object, method_name)
        # Call the method with arguments
        return method(*args)
    except AttributeError:
        print(f"Method '{method_name}' does not exist.")
    except Exception as e:
        print(f"Error invoking method '{method_name}': {e}")


def getDataFromEndOfLastBlock(doc, channelNumber, blockNumber, duration):
    lastBlock = doc.NumberOfRecords -1
    # numberOfBlocks = doc.
    secsPerTick = doc.GetRecordSecsPerTick(lastBlock)
    ticksPerSecond = 1/secsPerTick
    n_samples = round(duration*ticksPerSecond)
    endOfBlockTick = doc.GetRecordLength(lastBlock) # Get the last block
    start_sample = endOfBlockTick - n_samples
    data = doc.GetChannelData(0,channelNumber,lastBlock + 1,start_sample,n_samples)
    return data

# ------------------------------------------------------------------------------

# Connect to 
labchart = win32com.client.Dispatch("ADIChart.Application") 
filepath = "C:/Users/jeverts/Documents/Code/python_labchart_macros/test_file.adicht"                   # Enter the path to your labchart file e.g "C:/Users/yourname/Documents/your_file.adicht"
doc = labchart.Open(filepath)   # Open the specified document in Labchart   


# User Inputs ------------------------------------------------------------------
channelForGuideLine = 0        # The channel you want to add guidelines to
percentageOfMax = 20	        # The percentage of the maximum you want the second guideline to be displayed at
maximumColumnInDataPad = 1      # The column used on the datapad for keeping track of the maximum
waitTimeBeforeRecalculate = 4   # The time to wait before recalculating the maximum
# ------------------------------------------------------------------------------

rollingMaximum = 0		        # Variable to hold the rolling maximum value


# Setup the two guidelines
invoke_com_method(doc, "SetGuidelineValue", channelForGuideLine , 1, 0, "V", "")
invoke_com_method(doc, "SetGuidelineValue", channelForGuideLine , 2, 0, "V", "")
	
# Begin SetGuidelinesInfo
Guideline = 1
EnableGuideline = True
ShowGuideline = True
GuidelineColor = 9013641
# doc.SetGuidelinesInfo(channelForGuideLine, Guideline, EnableGuideline, ShowGuideline, GuidelineColor)
invoke_com_method(doc, "SetGuidelinesInfo", channelForGuideLine, Guideline, EnableGuideline, ShowGuideline, GuidelineColor)
# End SetGuidelinesInfo

# Begin SetGuidelinesInfo
Guideline = 2
EnableGuideline = True
ShowGuideline = True
GuidelineColor = 9013641
invoke_com_method(doc, "SetGuidelinesInfo", channelForGuideLine, Guideline, EnableGuideline, ShowGuideline, GuidelineColor)
# End SetGuidelinesInfo
	
# Begin SetGuidelineRegionInfo
ShadeTop = False
ShadeTopColor = 12763902
ShadeMid = True
ShadeMidColor = 12975793
ShadeBottom = False
ShadeBottomColor = 12763902
invoke_com_method(doc, "SetGuidelineRegionInfo", channelForGuideLine, ShadeTop, ShadeTopColor, ShadeMid, ShadeMidColor, ShadeBottom, ShadeBottomColor)
# End SetGuidelineRegionInfo


invoke_com_method(doc, "SetGuidelineValue", channelForGuideLine, 1, rollingMaximum, "V", "")
invoke_com_method(doc, "SetGuideLineValue", channelForGuideLine, 2, rollingMaximum - percentageOfMax/100*rollingMaximum , "V", "")


# Loop to calculate the maximum every 
while True:
    time.sleep(waitTimeBeforeRecalculate)
    block_number = 1
    start_sample = 1
    n_samples = 10
    chan_number = 1
    start_time = 0
    data = getDataFromEndOfLastBlock(doc, chan_number,block_number,waitTimeBeforeRecalculate)
    maxValueInMostRecentSelection = max(data)
    print("Rolling maximum:", rollingMaximum)
    print("Max in most recent data:", maxValueInMostRecentSelection, "\n")

    if (maxValueInMostRecentSelection > rollingMaximum):
        rollingMaximum = maxValueInMostRecentSelection 


    invoke_com_method(doc, "SetGuidelineValue", channelForGuideLine, 1, rollingMaximum, "V", "")
    invoke_com_method(doc, "SetGuideLineValue", channelForGuideLine, 2, rollingMaximum - percentageOfMax/100*rollingMaximum , "V", "")

	



