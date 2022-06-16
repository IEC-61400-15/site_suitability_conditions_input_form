# IEC 61400-15-1 Site Suitability Conditions Input Form
The IEC 61400-15-1 standard defines a framework for assessment and reporting of the wind turbine suitability conditions for both onshore and offshore power plants.This repository contains the human and machine readable data formats and documentation that support the Internatinal Electro_technical Commitee (IEC) 61400-15-1 site suitability conditions input form. For a complete copy of the IEC 61400-15-1 standard please contact the IEC at (https://www.iec.ch/homepage)


Wind turbine manufacturers require site-specific climatic conditions to evaluate the suitability (both fatigue and extreme loads) of wind turbines proposed for use at projects under development. Climatic conditions related to the wind regime are quantified using data collected by means of sensors and measurement devices, broadly including anemometry (cup and sonic) and remote sensing hardware (LiDAR and SODAR). Temperature and barometric pressure data are also required, and therefore often collected on-site as well. Data used for the quantifying climatic conditions may arrive in a variety of formats including (but not limited to), raw binary data, exported text values, or spreadsheet format. This repository includes two representative data formats which will serve as reference implementations for data exchange.

The meteorological and wind flow characteristics addressed in this part of the IEC 61400 relate to wind conditions, where parameters such as wind speed, wind direction, turbulence intensity, wind shear, inflow angle, air density and air temperature are included to the extent that they affect the operation and structural integrity of wind turbines. The period of record for these data may span six months to several years.
Every turbine manufacturer has a preferred format for receiving data and metadata for site suitability assessment and related data summaries. Additionally, energy developers and consultants have distinct preferred methods for presenting data and summary information. Although some preferred formats are similar, no two are exactly alike.

This framework provides for organization of wind and other climatic condition parameters for individual wind turbine and measurement device locations, as identified by geographic position and height above ground level. Furthermore, this framework addresses documentation and reporting requirements to facilitate traceability of the site suitability assessment processes. Traceability aids interpretation of data, and understanding of analysis results by all wind project development stakeholders.

Furthermore, the framework presents a Digital Exchange Format, consisting of complementary excel and JSON files, embodying the above mentioned reporting requirements and data structure. The excel and JSON files contain identical information. The excel file is useful for human readability, while JSON format is better suited for machine readability. The Digital Exchange Format contains the minimum requirements of data for assessment of the suitability of a wind project development site for any contemplated turbine make/model, and together with the report provides a clear view of project meteorological conditions

## Usage
The .xlsx file and Javasript Object Notation (JSON) file are intended to faciliate and streamline data exchange between wind industry stakeholders.  The IEC team has also included a versions that have been populated with example data. 

### Excel Versions:
Template: https://github.com/IEC-61400-15/site_suitability_conditions_input_form/blob/main/Site%20Suitability%20Input%20Conditions%20Form%20v15(20220616)_with_example_data.xlsx

Template with example data: https://github.com/IEC-61400-15/site_suitability_conditions_input_form/blob/main/Site%20Suitability%20Input%20Conditions%20Form%20v15(20220616)_with_example_data.xlsx

### JSON Versions
Template:

Template with example data:

## Issues
If you find any bugs or inconisstencies in the guidance please post an issue in this repository here: https://github.com/IEC-61400-15/site_suitability_conditions_input_form/issues
