# This file contains an example script for controlling LabChart within a Python
# environment. 
# 
# This script generates two guidelines within the channel defined by `channelIndex`.
# As data is sampled, the first guideline dynamically adjusts to the maximum data 
# point (within the most recent block), while the second guideline remains fixed 
# at a user-defined percentage below the maximum.
#
# This script must be run after sampling has started.


import win32com.client
import time
import pythoncom

# ------------------------------------------------------------------------------

def invoke_com_method(com_object, function_name, *args):
    """
    Invokes a method on a COM object by its name.

    Parameters:
        com_object: The COM object (created using win32com.client.Dispatch).
        function_name (str): The name of the method to invoke.
        *args: Arguments to pass to the method.

    Returns:
        The result of the method call, or None if the method fails.
    """
    try:
        # Get the DISPID (Dispatch ID) of the function by name
        dispid = com_object._oleobj_.GetIDsOfNames(0, function_name)
        
        # Invoke the function using the DISPID
        result = com_object._oleobj_.Invoke(
            dispid,  # DISPID of the function
            0,  # Reserved, must be 0
            win32com.client.pythoncom.DISPATCH_METHOD,  # Indicates a method call
            1,  # Indicates arguments are being passed
            *args  # Arguments to pass to the function
        )
        return result
    except pythoncom.com_error as e:
        print(f"Failed to invoke function '{function_name}': {e}")
        return None


# Helper function
def getDataFromEndOfLastBlock(doc, channelNumber, duration):
    lastBlock = doc.NumberOfRecords -1
    secsPerTick = doc.GetRecordSecsPerTick(lastBlock)
    ticksPerSecond = 1/secsPerTick
    n_samples = round(duration*ticksPerSecond)
    endOfBlockTick = doc.GetRecordLength(lastBlock) # Get the last block
    start_sample = endOfBlockTick - n_samples
    data = doc.GetChannelData(0,channelNumber,lastBlock + 1,start_sample,n_samples)
    return data

# ------------------------------------------------------------------------------

# Connect to Labchart
labchart = win32com.client.Dispatch("ADIChart.Application") 
doc = labchart.ActiveDocument    


# User Inputs ------------------------------------------------------------------
channelForGuideLine = 0         # The channel you want to add guidelines to
percentageOfMax = 20	        # The percentage of the maximum you want the second guideline to be displayed at
waitTimeBeforeRecalculate = 2   # The time to wait (in seconds) before recalculating the maximum

# ------------------------------------------------------------------------------

# Variable to hold the rolling maximum value
rollingMaximum = 0		        

# Setup the initial value of guideline 1
guideLineNumber = 1
guideLineValue = 0
guideLineUnits = "V"
guideLinePrefix = ""
invoke_com_method(doc, "SetGuidelineValue", channelForGuideLine , guideLineNumber, guideLineValue, guideLineUnits, guideLinePrefix)

# Setup the initial value of guideline 2
guideLineNumber = 2
guideLineValue = 0
guideLineUnits = "V"
guideLinePrefix = ""
invoke_com_method(doc, "SetGuidelineValue", channelForGuideLine , guideLineNumber, guideLineValue,guideLineUnits, guideLinePrefix)
	
# Set guideline 1 to be visible
guidelineNumber = 1
enableGuideline = True
showGuideline = True
guidelineColor = 9013641
invoke_com_method(doc, "SetGuidelinesInfo", channelForGuideLine, guidelineNumber, enableGuideline, showGuideline, guidelineColor)

# Set guideline 2 to be visible
guidelineNumber = 2
enableGuideline = True
showGuideline = True
guidelineColor = 9013641
invoke_com_method(doc, "SetGuidelinesInfo", channelForGuideLine, guidelineNumber, enableGuideline, showGuideline, guidelineColor)
	
# Set guideline shading to be green
shadeTop = False
shadeTopColor = 12763902
shadeMid = True
shadeMidColor = 12975793
shadeBottom = False
shadeBottomColor = 12763902
invoke_com_method(doc, "SetGuidelineRegionInfo", channelForGuideLine, shadeTop, shadeTopColor, shadeMid, shadeMidColor, shadeBottom, shadeBottomColor)


# Loop to check for a new maximum   
while True:
    time.sleep(waitTimeBeforeRecalculate)
    data = getDataFromEndOfLastBlock(doc, channelForGuideLine + 1,waitTimeBeforeRecalculate)
    maxValueInMostRecentSelection = max(data)

    if (maxValueInMostRecentSelection > rollingMaximum):
        rollingMaximum = maxValueInMostRecentSelection 

    # Update the guidelines with the latest maximum value
    invoke_com_method(doc, "SetGuidelineValue", channelForGuideLine, 1, rollingMaximum, "V", "")
    invoke_com_method(doc, "SetGuideLineValue", channelForGuideLine, 2, rollingMaximum - percentageOfMax/100*rollingMaximum , "V", "")



