# Excel2STIX

## General Information:
The purpose of the excel2stix.py Python-based script is to generate a STIX XML output file from a Microsoft Excel spreadsheet. The script excel2stix.py uses several Python packages that must be installed before use:  
* __lxml__ - Binds python to the libxml2 and libxslt libraries used for xml analysis 
* __openpyxl__ - Used to read and write Excel 2010 and later xlsx/xlsm/xltx/xltm files  
* __stix__ - Provides an API for creating and/or processing STIX content 
* __cybox__ - Used to parse, manipulate, and generate CybOX content 
* __jdcal__ - Contains functions used to convert between Julian dates and calendar dates 

The Microsoft Excel spreadsheet passed as an input to the excel2stix.py script must conform to certain standards to be recognized by excel2stix.py. Said standards regarding the necessary format of the input Excel spreadsheet can be found within the template.xlsx file provided with the script. If the file was not provided with the script, contact either the supplier of the script or the US-Cert software development team (contact information provided below). The output STIX XML file conforms to STIX version 1.1.n. The excel2stix.py script is capable of running from most Operating Systems to include Windows, Apple, and Linux variants. Easiest method of execution is through the operating system’s “cmd” command prompt.

## Pyton Library Dependencies:

The Excel2STIX.py script utilizes multiple python libraries to handle certain tasks within the process of converting from spreadsheet format to STIX format. These libraries are not included with any version of python and may need to be downloaded separately. The list of needed python libraries is stated above. A full list of the python libraries needed to be imported includes:

* __Import os__ – Provides operating system dependent functionality
* __Import sys__ – Provides access to variables and functions used by the python interpreter 
* __Import openpyxl__ - Used to read and write Excel 2010 and later xlsx/xlsm/xltx/xltm files 
* __Import pprint__ – stands for “pretty-print” and is used to print arbitrary python data structures 
* __Import time__ – Provides various date and time related functions 
* __Import uuid__ – (Universal Unique Identifier) Used to generate unique IDs 
* __Import codecs__ – Defines base classes for encoders and decoders and provides access to the internal Python codec registry 
* __Import warnings__ – Shows all necessary warning messages 
* __Import stix.common__ – Includes basic functions for creating and parsing STIX content 
  * __Import stix.common.kill_chains__
  * __Import stix.core__
  * __Import stix.indicator__
* __Import cybox__ - Used to parse, manipulate, and generate CybOX content 
* From stix.common.vocabs import vacabString – Defines controlled vocabulary implementations

In order to run excel2STIX.py, first check the system to see if the dependent Python packages have been installed. Installation is performed from the Command Prompt window.  There is no need to have administrator level access to your system. First run Python.  If the Python prompt doesn’t appear, either Python is not installed, or the environmental path is not set to the Python install folder.  Python is typically installed to __c:\Python27__ on Windows. Display the Windows environmental path by typing the command __echo %PATH%__ in a Command Prompt window. This environmental path can be augmented on Windows from the Control Panel, see System Properties; Advanced.

From the Command Prompt window, determine if the necessary Python packages have already been installed.  For example, figure 1 below shows how to check the versions of lxml, jdcal, stix, and cybox. Install a Python package if an error is received after attempting one of the import statements in figure 1 (see the “Installing Packages” section below).

Several Python packages are included in the downloads folder of this project as needed to install missing components.  Optionally, download sites are provided below to download the Python packages directly:

##### Table 1 - Download Sites

| Python Package | Download Site |
| -------------- | ------------- |
| Jdcal 1.0 | https://pypi.python.org/pypi/jdcal |
| Openpyxl 2.3.0 | https://pypi.python.org/pypi/openpyxl | 
| Lxml 3.4.0 | https://pypi.python.org/pypi/lxml |
| Stix 1.1.1.3 | https://pypi.python.org/pypi/stix |
| Cybox 2.1.0.9 | https://pypi.python.org/pypi/cybox/2.1.0.12 |

![][fig1]
##### Figure 1 – Checking Versions 

## Installing Packages:

Installing Python packages is pretty easy.  To install the Openpyxl package, step into the ~downloads/openpyxl-2.3.0 directory in a Command Prompt window and issue the command __setup.py install__.  Most of the Python packages support this mechanism for installation.  Another way to install these Python packages is to find the “dist” folder within the package (for example __~downloads/openpyxl-2.3.0/dist__) and copy the “.egg” folder to the python sitepackages folder (for example __c:/Python27/lib/site-packages/openpyxl-2.3.0-py2.7.egg__).  The lxml Python package must be at least version 3.3.1.  Table 1 lists all compatible versions used in the testing of the excel2stix.py script. When you installed everything, you should see the following folders under c:\Python27\Lib\site-packages:
* cybox-2.1.0.9-py2.7.egg 
*  jdcal-1.0-py2.7.egg 
* openpyxl-2.3.0-py2.7.egg 
* stix-1.1.1.3-py2.7.egg 
* lxml-3.4.3-py2.7-win32.egg.  

## Excel Template:  
A Microsoft Excel spreadsheet template is provided with the project.  See figure 2 below. Note that there are eleven worksheets:  Main, URL, IPv4, Link, File, E-mail, User Agent, Mutex, Registry, FQDN, and Network Connection.  Ten of these worksheets mirror the Indicators used by the IBTool.  The Main worksheet mirrors the overall IBTool metadata page.  Each of the individual indicator worksheets also mirrors each of the attributes in the IBTool for that specific Indicator.  Some data validation has also been added to the Excel template in the form of pulldowns for specific columns which should assist in populating data cells.  Remove unnecessary worksheets.  __Do not change the names of the worksheet tabs__, the excel2stix.py script looks specifically for these names. The Python script also expects to have the first row of labels (e.g., Description, Type, TLP, FOUO, etc) so __do not omit this row__.  

Within the “Main” tab of the template, two additional columns have been added to give the user the option to set the “Namespace_URL” and “Namespace_Tag” variables within the script. Simply change any of the desired data points for each column (e.g. the date, title of the document, TLP color, the namespace URL, ect.) and save the excel document, preferably in the same directory as the script.

![][fig2]
##### Figure 2 – Excel Template

## Excel2STIX Execution:

The Python script excel2stix.py is pretty easy to use.  This script expects only one command line argument:  an Excel spreadsheet which conforms to the given template.  It takes a minute or two to read the Microsoft Excel spreadsheet from within the script, so __don’t worry that the script is hanging or has crashed__.  The output file will be something like 20151103_102236.xml, where 20151103 is the year, month, and day and the 102236 is the hour, minute, and second when the file was created.  The output file will conform to STIX version 1.1.1.  Figure 3 shows what a typical session should look like. 

![][fig3]
##### Figure 3 – Typical Script Session

# CYBERSECURITY AND INFRASTRUCTURE SECURITY (CISA) TERMS OF USE FOR EXCEL2STIX, VERSION 1.0a

## DEVELOPMENT OF EXCEL2STIX.
The Cybersecurity and Infrastructure Agency (CISA) of the U.S. Department of Homeland Security leads a collaborative effort with industry to develop a standardized, structured language to represent cyber threat information called the “Structured Threat Information eXpression” (STIX™). The STIX™ framework intends to convey the full range of potential cyber-threat-data elements and strives to be as expressive, flexible, extensible, automatable, and human-readable as possible. To facilitate the use of the STIX™ framework to represent cyber threat information, CISA developed an excel2stix.py Python-based script to generate STIX™ XML output file from a Microsoft Excel spreadsheet. Excel2STIX 1.0a allows users who have collected cyber threat information in a Microsoft Excel spreadsheet to convert that information into the STIX™ framework.

## WHAT IS EXCEL2STIX?
(a) Excel2STIX 1.0a is python-based script that can generate a STIX™ XML output file from a Microsoft Excel spreadsheet. The script excel2stix.py uses several python packages that must be installed before use, including:
* Lxml, which binds python to the libxml2 and libxslt libraries used for xml analysis
* Openpyxl, which is used to read and write Excel 2010 and later xlsx/xlsm/xltx/xltm files
* Stix, which provides an API for creating and/or processing STIX™ content
* Cybox, which is used to parse, manipulate, and generate CybOX™ content
* Jdcal, which contains functions used to convert between Julian dates and calendar dates

(b) To convert the data in the Microsoft Excel spreadsheet to STIX™ XML, the Microsoft Excel Spreadsheet must conform to certain standards. CISA is therefore providing the appropriately formatted Microsoft Excel spreadsheet in template.xlsx format with the excel2stix.py script. The output STIX™ XML file generated from the Microsoft Excel spreadsheet conforms to STIX™ version 1.1.n. The excel2stix.py script is capable of running from most operating systems, including Windows, Apple, and Linux variants. The Excel2STIX User Guide contains detailed information about installing and operating the excel2stix.py script.

## OPEN SOURCE DISTRIBUTION OF EXCEL2STIX.
CISA desires to distribute the excel2stix.py script to the public through an open source platform to allow users to have access to the script in source code and binary format. CISA also believes the open source distribution of Excel2STIX 1.0a will make it easier for adoption of the STIX™ format and allow stakeholders to share cyber threat information with CISA that conforms to STIX™ version 1.1.n. Accordingly, CISA provides Excel2STIX 1.0a to you, the user, under the conditions contained in this Terms of Use.
## LICENSE GRANTED IN EXCEL2STIX.
A U.S. Government contractor developed Excel2STIX 1.0a for the Cybersecurity and Infrastructure Security Agency of the U.S. Department of Homeland Security and therefore Excel2STIX 1.0a is subject to United States copyright law. The United States Government has unlimited rights in the copyright in Excel2STIX 1.0a, which is sufficient to allow end users to download, access, install, copy, modify, and otherwise use Excel2STIX 1.0a for its intended purpose. Specifically, the U.S. Government is providing Excel2STIX 1.0a to Users with a royalty-free, irrevocable, worldwide license to use, disclose, reproduce, prepare derivative works, distribute copies to the public, including by electronic means, and perform publicly and display publicly Excel2STIX 1.0a, in any manner, including by electronic means, and for any purpose whatsoever.

## ADDITIONAL LICENSE TERMS.
(a) User will ensure compliance with, and provide continued notice of this license, and any open sources licenses identified in Section 6 of this Terms of Use.

(b) To the extent practicable, User will acknowledge the contribution of DHS/CISA in the development of the Excel2STIX 1.0a by the following statement in a readme text file of the tool: Excel2STIX 1.0a is developed with funds from the Cyber Security and Infrastructure Security Agency of the U.S. Department of Homeland Security.

(c) This Terms of Use does not constitute, in any matter, an endorsement by CISA, DHS, or the U.S. Government of any information, plans, or actions resulting from use of Excel2STIX 1.0a.

(d) This Terms of Use does not, in any manner, constitute the grant of a license to the public of any CISA or DHS, or third party patent, patent application, copyright or trademark, except as provided by this Terms of Use, or other intellectual property of CISA and DHS.

(e) CISA denies any, and all, liability, as discussed in Section 7 of this Terms of Use, associated with or resulting from the use of Excel2STIX 1.0a

(f) This Terms of Use and the legal relations between the Users and CISA shall be determined in accordance with United States Federal law.

## NOTICE OF THIRD PARTY SOFTWARE.
(a) The Excel2STIX 1.0a does not contain proprietary or open source software. However, the excel2stix.py script utilizes multiple python libraries to handle certain tasks within the process of converting from spreadsheet format to STIX™ format. These libraries are not included with any version of python and User may need to download separately. The list of python libraries needed to be imported is in Attachment A of this Terms of Use. All components of the Excel2STIX 1.0a, individually and as a combined work are subject to United States Copyright law.

(b) All third party software necessary for operation of the Excel2STIX 1.0a are subject to copyright licenses. DHS has identified all the third party copyright licenses needed to operate Excel2STIX 1.0a and conducted a good faith analysis to determine that the third party library dependencies allow users to download, access, install, copy, modify, distribute and otherwise use the third party library dependencies for operation of Excel2STIX 1.0a.

(c) Identified in Attachment A, are the third party library dependencies and links to the specific terms and conditions, not under CISA or DHS control, but that are applicable to the above referenced third party software.

## DISCLAIMER OF LIABILITY.
The United States Government shall be not be liable or responsible for any maintenance or updating of Excel2STIX 1.0a, nor for correction of any errors in Excel2STIX 1.0a.

THE EXCEL2STIX 1.0a IS PROVIDED “AS IS” WITHOUT ANY WARRANTY OF ANY KIND, EITHER EXPRESSED, IMPLIED, OR STATUTORY, INCLUDING, BUT NOT LIMITED TO, ANY WARRANTY THAT THE EXCEL2STIX 1.0a WILL CONFORM TO SPECIFICATIONS, ANY IMPLIED WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR FREEDOM FROM INFRINGEMENT, ANY WARRANTY THAT THE EXCEL2STIX 1.0a WILL BE ERROR FREE. IN NO EVENT SHALL THE UNITED STATES GOVERNMENT OR ITS CONTRACTORS OR SUBCONTRACTORS BE LIABLE FOR ANY DAMAGES, INCLUDING, BUT NOT LIMITED TO, DIRECT, INDIRECT, SPECIAL OR CONSEQUENTIAL DAMAGES, ARISING OUT OF, RESULTING FROM, OR IN ANY WAY CONNECTED WITH THE EXCEL2STIX 1.0a, WHETHER OR NOT BASED UPON WARRANTY, CONTRACT, TORT, OR OTHERWISE, WHETHER OR NOT INJURY WAS SUSTAINED BY PERSONS OR PROPERTY OR OTHERWISE, AND WHETHER OR NOT LOSS WAS SUSTAINED FROM, OR AROSE OUT OF THE RESULTS OF, OR USE OF, THE EXCEL2STIX 1.0a. THE UNITED STATES GOVERNMENT DISCLAIMS ALL WARRANTIES AND LIABILITIES REGARDING THIRD PARTY SOFTWARE, IF PRESENT IN THE EXCEL2STIX 1.0a, AND DISTRIBUTES IT “AS IS”.

### ATTACHMENT A
### Python Library Dependencies

### BSD License For:
Jdcal, Version=1.0: Contains functions used to convert between Julian dates and calendar dates https://pypi.python.org/pypi/jdcal; https://opensource.org/licenses/BSD-3-Clause

Lxml, Version=3.4.0: Binds python to the libxm12 and libxslt libraries used for xml analysis
https://pypi.python.org/pypi/lxml; https://opensource.org/licenses/BSD-3-Clause

Stix, Version=1.1.1.3: Provides an API for creating and/or processing STIX™ content
https://pypi.python.org/pypi/stix; https://opensource.org/licenses/BSD-3-Clause

Cybox, Version=2.1.0.9: Used to parse, manipulate, and generate CybOX content
https://pypi.python.org/pypi/cybox/2.1.0.12; https://opensource.org/licenses/BSD-3-Clause

### MIT License For:
Openpyxl Version=2.3.0: Used to read and write Excel 2010 and later xlsx/xlsm/sltx/xltm files
https://pypi.python.org/pypi/openpyxl; https://opensource.org/licenses/MIT

### End of Python Library Dependencies Legal information

[fig1]: img/fig1.png "Figure 1 - Checking Versions"
[fig2]: img/fig2.png "Figure 2 - Excel Template"
[fig3]: img/fig3.PNG "Figure 3 - Typical Script Session"
