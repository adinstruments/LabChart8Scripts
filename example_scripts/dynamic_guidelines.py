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
filepath = "C:/Users/jeverts/Documents/Code/python_labchart_macros/test_file.adicht" # Enter the path to your labchart file e.g "C:/Users/yourname/Documents/your_file.adicht"
doc = labchart.Open(filepath)   # Open the specified document in Labchart   


# User Inputs ------------------------------------------------------------------
channelForGuideLine = 0         # The channel you want to add guidelines to
percentageOfMax = 20	        # The percentage of the maximum you want the second guideline to be displayed at
waitTimeBeforeRecalculate = 2   # The time to wait before recalculating the maximum

# ------------------------------------------------------------------------------

# Variable to hold the rolling maximum value
rollingMaximum = 0		        

# Setup the two guidelines
invoke_com_method(doc, "SetGuidelineValue", channelForGuideLine , 1, 0, "V", "")
invoke_com_method(doc, "SetGuidelineValue", channelForGuideLine , 2, 0, "V", "")
	
# Set guideline 1 to be visible
Guideline = 1
EnableGuideline = True
ShowGuideline = True
GuidelineColor = 9013641
invoke_com_method(doc, "SetGuidelinesInfo", channelForGuideLine, Guideline, EnableGuideline, ShowGuideline, GuidelineColor)

# Set guideline 2 to be visible
Guideline = 2
EnableGuideline = True
ShowGuideline = True
GuidelineColor = 9013641
invoke_com_method(doc, "SetGuidelinesInfo", channelForGuideLine, Guideline, EnableGuideline, ShowGuideline, GuidelineColor)
	
# Set guideline shading to be green
ShadeTop = False
ShadeTopColor = 12763902
ShadeMid = True
ShadeMidColor = 12975793
ShadeBottom = False
ShadeBottomColor = 12763902
invoke_com_method(doc, "SetGuidelineRegionInfo", channelForGuideLine, ShadeTop, ShadeTopColor, ShadeMid, ShadeMidColor, ShadeBottom, ShadeBottomColor)

# Initialise the guidelines to 0
invoke_com_method(doc, "SetGuidelineValue", channelForGuideLine, 1, 0, "V", "")
invoke_com_method(doc, "SetGuideLineValue", channelForGuideLine, 2, 0 , "V", "")


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



