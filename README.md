# HP 3456A Control and Data Logging Software
This software allows you to control and log data from the HP 3456A multimeter via an AR488 Arduino GPIB Adapter. Supports Window 10, 8, and 7.

#### [Download](https://github.com/Niravk1997/HP-3456A-Software/releases)

#### [EEVblog forum post](https://www.eevblog.com/forum/testgear/hp-34401a-standalone-software/)

#### [AR488 Adapter](https://github.com/Twilight-Logic/AR488)



#### The main software window
![HP 3456A Software](https://github.com/Niravk1997/HP-3456A-Software/blob/main/Images/HP3456A_Main_Window.gif)

#### Interactive Graphing Module
![HP 3456A Graph Module](https://github.com/Niravk1997/HP-3456A-Software/blob/main/Images/HP3456A_Graph_Module.gif)

**Features:**
- Control and Log data, save data into organized folders
- Multithreading Support:
   - All Windows open in a new thread, this ensures maximum performance. For example: Interacting with the Graph Window does not slow down other Windows.
    - All Serial communication happens on an excusive thread. This allows the software to maintain maximum sample capture speed at all times, regardless of what other data processing might be going on.
    - Users can interact with the Graph Window smoothly without any lag.
    - Data Logging functions also run on their designated thread, periodically the software will save measurement data from FIFO data structures into text files.
- Speech Synthesizer feature allows the software to voice measurements periodically and or when it meets the maximum or minimum value threshold.
- Graph Window allows users to visualize their captured data. You can get statistics for all the samples capture or for select few samples. Pan, zoom, and zoom to highlighted area. Save/copy graph as image or save graph's data into text/csv files. 
- Create math waveforms from the samples captured. Create math waveforms from math waveforms. There is no limit to how many math waveforms you can create. 
- Create Histogram from the samples captured and from math waveforms. There is no limit to how many histogram waveforms you can create.
- Measurement table allows users to collect and display the measurement data into a table.
- Capture up to 310  measurement samples in 2 seconds. Fastest sample capture rate compared to any other HP 3456A GUI software.

#### Create Math Waveform
![HP 3456A Math Waveform](https://github.com/Niravk1997/HP-3456A-Software/blob/main/Images/HP3456A_Math_Waveform.gif)

#### Histogram Waveform
![HP 3456A Histogram](https://github.com/Niravk1997/HP-3456A-Software/blob/main/Images/HP3456A_Histogram_Window.gif)

#### Measurement Table
![HP 3456A Table](https://github.com/Niravk1997/HP-3456A-Software/blob/main/Images/HP3456A_Table.gif)