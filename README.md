# json2xlsx
  A Generic Json to Excel Xlsx Sheet Convert Util

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

![image](https://github.com/user-attachments/assets/fbdb789f-5140-49be-b30b-da935cf8401c)




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
✗ python3 json2xlsx.py input.json output.xlsx --dictKeyHeader=Predefined-Header-Col
![image](https://github.com/user-attachments/assets/ad84786e-d38f-451b-81b8-ae4dc4796130)

# Limitation:
If there are multiple lists exist in input json, they must be homogeneous in structures.
A counter exmaple is:
```json
{
  "list-1st": [
    {
      "A": "VAL-A1",
      "B": "VAL-B1"
    },
    {
      "A": "VAL-A2",
      "B": "VAL-B2"
    }
  ],
  "list-2nd": [
    {
      "C": "VAL-C1",
      "D": "VAL-D1"
    },
    {
      "C": "VAL-C2",
      "D": "VAL-D2"
    }
  ]
}
```



