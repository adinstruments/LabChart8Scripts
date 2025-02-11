# LabChart 8: Python Template Scripts

[LabChart 8](https://adi.to/labchart)
is (//TODO: Generic sales pitch...) from
ADInstruments. This repository contains a number of templates that demonstrate
how to programmatically control the Labchart application using a Python script.

<!-- ---

The repository contains:
 + [`stimulator_scripts/`](./stimulator_scripts): Examples of how to use
   LabChart Lightning's scripting capabilities to output custom stimulus
   waveforms.
 + [`table_analyses/`](./table_analyses): Template scripts for various
   programming languages that show how to run statistical analyses on the
   summary data exported from LabChart Lightning. -->



## Automation using COM

The Labchart application can be controlled programmatically using **Windows COM (Component Object Model)**, a software communication technology created by Microsoft. 

Labchart provides an API that enables interaction with the application through a COM interface. This interface can be accessed in a Python script using the `win32com.client` module (a module from the `pywin32` library, which offers Python bindings for the Windows COM API).

A simple example is shown below:

```python
import win32com.client

## Labchart must be open before running this script.

# Instantiate the labchart COM object.
labchart = win32com.client.Dispatch("ADIChart.Application") 

# Enter the path to your labchart file e.g "C:/Users/yourname/Documents/your_file.adicht"
filepath = "" 

# Open the specified document in Labchart   
doc = labchart.Open(filepath)

# Add the comment "Hello World!" to the end of channel 1
channelIndex = 0 # Note: Labchart lanes are zero indexed. 
doc.AddCommentAtEnd(channelIndex, "Hello World!")
```

The `doc` object instantiated in the code example above can be used to an array of call commands specific to the opened document. A list of functions available on the document object can be found in the Macro Editor within the Labchart application (accesible by either creating a new macro or editing an existing macro).

(//TODO: Improve quality of screenshot below and add annotations)

![Example Image](macro_editor.png)





---
