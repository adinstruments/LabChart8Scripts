# Event Listeners Template

`event_listeners.py` is a template script for listening and responding to LabChart events within a Python environment. 

Events available to listen to are:

- "OnStartSampling", event fires when the sampling session is started (Start button pushed).

- "OnStartSamplingBlock", event fires when sampling and a new block is about to be added to the document. This is called *before* the new block has been added, so `document.NumberOfRecords` does not yet include the new block.

- "OnNewSamples", event fires (roughly 20 times/s) when new samples may be available".  
  
  Event returns:  
`args[0]` = The number of new samples added

- "OnFinishSamplingBlock", event fires when a sampling block is ended.

- "OnFinishSampling", event fires when a sampling session is ended.

- "OnSelectionChange", event fires when the selection changes.

- "OnCommentAdded", event fires when a comment (or comments) are added.  
  
  Event returns:  
`args[0]` = Comment text  
`args[1]` = Channel index comment was placed in  
`args[2]` = Block index comment was placed in  
`args[3]` = Tick index comment was placed


