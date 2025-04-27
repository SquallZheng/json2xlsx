# json-to-excel
  A Generic Json to Excel Xlsx Sheet Converte Util

# Usage:
python3 json2xlsx.py <input.json> <output.xlsx> [notMergeCell] [--dictKeyHeader]
  
```sh
➜  Scripts-Synthesis git:(master) ✗ python3 json2xlsx.py --help
usage: json2xlsx.py [-h] [--dictKeyHeader DICTKEYHEADER] jsonFile outputFileName

json2xlsx commands

positional arguments:
  jsonFile              Path of input json file
  outputFileName        Path of output xlsx file
  notMergeCell          False to prohibit merge continuous identical cell vertically

options:
  -h, --help            show this help message and exit
  --dictKeyHeader DICTKEYHEADER
                        If input json is a dictionary, extract its key as content of specified header column
```

# Examples:
```json
{
  "organization": "freeCodeCamp",
  "website": "https://www.freecodecamp.org/",
  "founder": "Quincy Larson",
  "a": {
    "b": {
      "c": "C"
    }
  },
  "certifications": [
    {
      "name": "Responsive Web Design",
      "courses": [
        "HTML",
        "CSS"
      ]
    },
    {
      "name": "JavaScript Algorithms and Data Structures",
      "courses": [
        "JavaScript"
      ]
    },
    {
      "name": "Front End Development Libraries",
      "courses": [
        "Bootstrap",
        "jQuery",
        "Sass",
        "React",
        "Redux"
      ]
    },
    {
      "name": "Data Visualization",
      "courses": {
        "name": "D3"
      }
    },
    {
      "name": "Relational Database Course",
      "courses": [
        "Linux",
        "SQL",
        "PostgreSQL",
        "Bash Scripting",
        "Git and GitHub",
        "Nano"
      ]
    },
    {
      "name": "Back End Development and APIs",
      "courses": [
        "MongoDB",
        "Express",
        "Node",
        "NPM"
      ]
    },
    {
      "name": "Quality Assurance",
      "courses": [
        "Testing with Chai",
        "Express",
        "Node"
      ]
    },
    {
      "name": "Scientific Computing with Python",
      "courses": [
        "Python"
      ]
    },
    {
      "name": "Data Analysis with Python",
      "courses": [
        "Numpy",
        "Pandas",
        "Matplotlib",
        "Seaborn"
      ]
    },
    {
      "name": "Information Security",
      "courses": [
        "HelmetJS"
      ]
    },
    {
      "name": "Machine Learning with Python",
      "courses": [
        "Machine Learning",
        "TensorFlow"
      ]
    }
  ],
  "formed": 2014
}
```
✗ python3 json2xlsx.py test.json test.json.xlsx

![image](https://github.com/user-attachments/assets/a903363a-14b2-4ad9-84ba-abb091bcbd65)




```json
{
  "H1": {
    "content": [
      {
        "KEY": "a"
      }
    ]
  },
  "H2": {
    "content": [
      {
        "KEY": "b"
      }
    ]
  },
  "H3": {
    "content": [
      {
        "KEY": "c"
      }
    ]
  }
}
```
✗ python3 json2xlsx.py input.json output.xlsx --dictKeyHeader=Predeined-Header-Col
![image](https://github.com/user-attachments/assets/e30885b8-bf93-4127-b334-38f83eb0f694)

# Limitation:
If there are multiple lists exist in input json, they must be homogeneous in structures.



