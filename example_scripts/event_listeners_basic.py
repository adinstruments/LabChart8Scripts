import win32com.client
import pythoncom

# ------------------------------------------------------------------------------

class LabChartEventHandler:
    def OnStartSamplingBlock(self, *args):
        """
        Example event handler called when sampling and a new block is about to be added to the document.
        This is called *before* the new block has been added, so document.NumberOfRecords does not yet include the new block.
        """
        print("OnBlockStart called")
    

    def OnNewSamples(self, *args):
        """
        Example event handler called when sampling, whenever new samples might be available,
        typically 20 times a second.
        This example gets any new samples from channel gChan, appends them to the gChan1Data list,
        then plots the latest 5000 samples.
        """
        print("OnNewSamples called")

    def OnBlockFinish(self, *args):
        """
        Example event handler called when a block finishes.
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