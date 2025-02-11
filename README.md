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


Labchart provides an API that enables interaction with the application through a **Windows COM (Component Object Model)** interface. This interface can be accessed in a Python script using the `win32com.client` module (*a module from the `pywin32` library that offers Python bindings for the Windows COM API*).

A simple example is shown below:

```python
import win32com.client

## Labchart must be open before running this script.

# Instantiate the labchart COM object.
labchart = win32com.client.Dispatch("ADIChart.Application") 

# Enter the path to your labchart file e.g "C:/Users/yourname/Documents/your_file.adicht"
filepath = "" 

# Open the specified document in Labchart   
document = labchart.Open(filepath)

# Add the comment "Hello World!" to the end of channel 1
channelIndex = 0 # Note: Labchart lanes are zero indexed. 
document.AddCommentAtEnd(channelIndex, "Hello World!")
```

The `document` object instantiated in the code example above, provides an interface for controlling the opened Labchart document. A list of functions available on the interface can be found by running the `dir()` command:

```
dir(doc)

['Activate', 'AddCommentAtSelection', 'AddRef', 'AddToDataPad', 'AppendComment', 'AppendFile', 'AppendFileEx', 'Application', 'Close', 'CreatePlot', 'FullName', 'GetChannelData', 'GetChannelName', 'GetDataPadColumnChannel', 'GetDataPadColumnFuncName', 'GetDataPadColumnUnit', 'GetDataPadCurrentValue', 'GetDataPadValue', 'GetDigitalInputBit', 'GetDigitalInputState', 'GetDigitalOutputBit', 'GetDigitalOutputState', 'GetIDsOfNames', 'GetName', 'GetPlot', 'GetPlotId', 'GetRecordLength', 'GetRecordSecsPerTick', 'GetRecordStartDate', 'GetScopeChannelData', 'GetSelectedData', 'GetSelectedValue', 'GetTypeInfo', 'GetTypeInfoCount', 'GetUnits', 'GetViewPos', 'ImportMacros', 'Invoke', 'IsChannelSelected', 'IsRecordMode', 'IsSampling', 'Macros', 'MatLabPutChannelData', 'MatLabPutFullMatrix', 'Name', 'NumberOfChannels', 'NumberOfDisplayedChannels', 'NumberOfRecords', 'PLCDebugCommand', 'Parent', 'Path', 'PlayMacro', 'PlayMessage', 'Print', 'QueryInterface', 'RecordTimeToTickPosition', 'Release', 'ResetSelection', 'SamplingRecord', ...]
```



(//TODO: Improve quality of screenshot below and add annotations)

![Example Image](macro_editor.png)





---
