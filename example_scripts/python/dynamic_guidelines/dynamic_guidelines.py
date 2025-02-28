# This file contains an example script for listening and responding to LabChart 
# events within a Python environment. 

# This script generates two guidelines within the channel defined by `channelIndex`.
# As data is sampled, the first guideline dynamically adjusts to the maximum data 
# point (within the most recent block), while the second guideline remains fixed 
# at a user-defined percentage below the maximum.

# ------------------------------------------------------------------------------

import win32com.client
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



def checkForNewMaximum(document, numberOfTicks):
    """
    Helper function executed on each 'OnNewSamples' event.
    Determines the maximum value among the new samples and checks if it exceeds 
    the current block-wide maximum.
    If a new block-wide maximum is found, the guidelines are updated.
    """
    lastBlock = document.NumberOfRecords
    endOfBlockTick = document.GetRecordLength(lastBlock)
    start_sample = endOfBlockTick - numberOfTicks
    data = document.GetChannelData(0,channelForGuideLine+1, lastBlock, start_sample+1, numberOfTicks)
    if (len(data) == 0):
        return
    maxValueInMostRecentSelection = max(data)


    global blockWideMaximum
    if (maxValueInMostRecentSelection > blockWideMaximum):
        blockWideMaximum = maxValueInMostRecentSelection 
        # Update the guidelines with the latest maximum value
        invoke_com_method(document, "SetGuidelineValue", channelForGuideLine, 1, blockWideMaximum, "V", "")
        invoke_com_method(document, "SetGuideLineValue", channelForGuideLine, 2, blockWideMaximum - percentageOfMax/100*blockWideMaximum , "V", "")

    

def inititaliseGuideLines(document): 
    # Reset blockWideMaximum back to 0
    global blockWideMaximum
    blockWideMaximum = 0


    # Setup the initial value of guideline 1
    guideLineNumber = 1
    guideLineValue = 0
    guideLineUnits = "V"
    guideLinePrefix = ""
    invoke_com_method(document, "SetGuidelineValue", channelForGuideLine , guideLineNumber, guideLineValue, guideLineUnits, guideLinePrefix)

    # Setup the initial value of guideline 2
    guideLineNumber = 2
    guideLineValue = 0
    guideLineUnits = "V"
    guideLinePrefix = ""
    invoke_com_method(document, "SetGuidelineValue", channelForGuideLine , guideLineNumber, guideLineValue,guideLineUnits, guideLinePrefix)
        
    # Set guideline 1 to be visible
    guidelineNumber = 1
    enableGuideline = True
    showGuideline = True
    guidelineColor = 9013641
    invoke_com_method(document, "SetGuidelinesInfo", channelForGuideLine, guidelineNumber, enableGuideline, showGuideline, guidelineColor)

    # Set guideline 2 to be visible
    guidelineNumber = 2
    enableGuideline = True
    showGuideline = True
    guidelineColor = 9013641
    invoke_com_method(document, "SetGuidelinesInfo", channelForGuideLine, guidelineNumber, enableGuideline, showGuideline, guidelineColor)
        
    # Set guideline shading to be green
    shadeTop = False
    shadeTopColor = 12763902
    shadeMid = True
    shadeMidColor = 12975793
    shadeBottom = False
    shadeBottomColor = 12763902
    invoke_com_method(document, "SetGuidelineRegionInfo", channelForGuideLine, shadeTop, shadeTopColor, shadeMid, shadeMidColor, shadeBottom, shadeBottomColor)


# User Inputs ------------------------------------------------------------------

channelForGuideLine = 0         # The channel you want to add guidelines to
percentageOfMax = 20	        # The percentage of the maximum you want the second guideline to be displayed at

# ------------------------------------------------------------------------------

# Variable to hold the block-wide maximum value
blockWideMaximum = 0		      

# Connect to Labchart
labchart = win32com.client.Dispatch("ADIChart.Application") 
document = labchart.ActiveDocument

# Setup the event handlers
class LabChartEventHandler:
    def OnStartSampling(self):
        """
        Example event handler called when the sampling session is started (Start button pushed).
        """
        inititaliseGuideLines(document)
        

    def OnNewSamples(self, *args):
        """
        Example event handler called (roughly 20 times/s) when new samples may be available")
        """
        checkForNewMaximum(document, numberOfTicks = args[0])



# Register the event handlers
win32com.client.WithEvents(document, LabChartEventHandler)

print("Script running. Listening for LabChart events...")
pythoncom.PumpMessages()  # Keeps the COM event loop alive






