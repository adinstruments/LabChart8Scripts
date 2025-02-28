# Dynamic Guidelines Example

`dynamic_guidelines.py` is an example script for listening and responding to LabChart events within a Python environment. 

The script generates two guidelines within a user defined channel. As data is sampled, the first guideline dynamically adjusts to the maximum data point (within the most recent block), while the second guideline remains fixed 
at a user-defined percentage below the maximum. The guidelines are reset on each new block. 

To use:
- Open a labchart document
- Set the `channelIndex` variable within `dynamic_guidelines.py` to the channel you wish the guidelines to be displayed in.
- Run the script. The script will run until closed, listening and responding to Labchart events.  

![dynamic_guidelines script running in Labchart](./dynamic_guidelines_example.gif)

