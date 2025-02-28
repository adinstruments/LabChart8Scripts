# This file contains an example script for listening and responding to LabChart 
# Events within a Python environment. 

# ------------------------------------------------------------------------------

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
        
        Event returns:
        args[0] = Number of new samples added
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

# ------------------------------------------------------------------------------

# Connect to Labchart
labchart = win32com.client.Dispatch("ADIChart.Application") 
document = labchart.ActiveDocument

# Register the event handlers
win32com.client.WithEvents(document, LabChartEventHandler)

print("Script running. Listening for LabChart events...")
pythoncom.PumpMessages()  # Keeps the COM event loop alive

# ------------------------------------------------------------------------------