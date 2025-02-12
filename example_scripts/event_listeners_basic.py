import win32com.client
import pythoncom

# ------------------------------------------------------------------------------

class LabChartEventHandler:
    def OnStartSampling(self, *args):
        """
        Example event handler called when the sampling session is started (Start button pushed).
        """
        print("OnStartSampling called")
        
    def OnStartSamplingBlock(self, *args):
        """
        Example event handler called when sampling and a new block is about to be added to the document.
        This is called *before* the new block has been added, so document.NumberOfRecords does not yet include the new block.
        """
        print("OnBlockStart called")
    

    def OnNewSamples(self, *args):
        """
        Example event handler called (roughly 20 times/s) when new samples may be available")
        """
        print("OnNewSamples called")

    def OnFinishSamplingBlock(self, *args):
        """
        Example event handler called when a sampling block is ended.
        """
        print("OnBlockFinish called")

    
# ------------------------------------------------------------------------------

# Connect to Labchart
labchart = win32com.client.Dispatch("ADIChart.Application") 
document = labchart.ActiveDocument

# Register the event handlers
win32com.client.WithEvents(document, LabChartEventHandler)

# Keep the script running
print("Script is running. Press Ctrl+C to exit.")
try:
    pythoncom.PumpMessages()  # Keeps the COM event loop alive
except KeyboardInterrupt:
    print("Exiting script.")

# ------------------------------------------------------------------------------