import win32com.client
import pythoncom

# ------------------------------------------------------------------------------

class LabChartEventHandler:
    def OnStartSampling(self):
        """
        Example event handler called when the sampling session is started (Start button pushed).
        """
        print("OnStartSampling called")
        
    def OnStartSamplingBlock(self):
        """
        Example event handler called when sampling and a new block is about to be added to the document.
        This is called *before* the new block has been added, so document.NumberOfRecords does not yet include the new block.
        """
        print("OnBlockStart called")
    

    def OnNewSamples(self, *args):
        """
        Example event handler called (roughly 20 times/s) when new samples may be available")
        """
        print("OnNewSamples called", args)

    def OnFinishSamplingBlock(self):
        """
        Example event handler called when a sampling block is ended.
        """
        print("OnBlockFinish called")

    def OnFinishSampling(self):
        """
        Example event handler called when a sampling session is ended.
        """
        print("OnFinishSampling called")

    def OnSelectionChange(self):
        """
        Example event handler called when the selection changes.
        """
        print("OnSelectionChange called")

    def OnCommentAdded(self, *args):
        """
        Example event handler called when a comment (or comments) are added.

        Event returns:
        args[0] = Comment text
        args[1] = Channel index comment was placed in
        args[2] = Block index comment was placed in 
        args[3] = Tick index comment was placed
        """
        print("OnCommentAdded called", args)

    def OnDigitalInputChanged(self, *args):
        """
        Example event handler called when there's a change in the digital input while sampling.

        Event returns:
        args[0] = digitalByteOld
        args[1] = digitalByteNew
        args[2] = tick
        """
        print("OnDigitalInputChanged called", args)

    
    def OnGuidelineCrossed(self, *args):
        """
        Example event handler called when a guideline is crossed in a specified channel.

        Event returns:
        args[0] = channelNumber
        args[1] = guidelineNumber
        args[3] = isRising
        args[4] = position
        args[5] = guidelineValue
        args[6] = signalValue
        """
        print("OnGuidelineCrossed called", args)

    def OnKeysPressed(self, *args):
        """
        Example event handler called when a key is pressed.

        Event returns:
        args[0] = key
        args[1] = isControlKeyDown
        args[3] = isShiftKeyDown
        """
        print("OnKeysPressed called", args)

    def OnDataPadSelectionChanged(self, *args):
        """
        Example event handler called when the Data Pad selection changes.

        Event returns:
        args[0] = Datapad sheet index 
        args[1] = column
        args[3] = row
        args[4] = width
        args[5] = height
        """
        print("OnDataPadSelectionChanged called", args)

    def OnEventDataArrived(self, *args):
        """
        Example event handler called when an event from a calculation such as Cyclic Measurements arrives.

        Event returns:
        args[0] = channelNumber
        args[1] = isInternalDetectorChannel
        args[3] = eventValue
        args[4] = position
        """
        print("OnEventDataArrived called", args)


    
# ------------------------------------------------------------------------------

# Connect to Labchart
labchart = win32com.client.Dispatch("ADIChart.Application") 
document = labchart.ActiveDocument

# Register the event handlers
win32com.client.WithEvents(document, LabChartEventHandler)

print("Script running. Listening for LabChart events...")
pythoncom.PumpMessages()  # Keeps the COM event loop alive

# ------------------------------------------------------------------------------