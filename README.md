# JAP_LLM_Platform

This project aims to build a Japanese self-study platform powered by Large Language Models (LLM). The main function of the platform is to store, mark and generate Japanese N4/N5 level question papers and provide users with personalized feedback.

## Features

- Generate Japanese N4/N5 level test questions using LLM
- Store and manage question bank in SQL database 
- Mark student test papers automatically
- Track student performance and provide feedback
- User-friendly graphical interface

## Project Structure

```
├── Source Code/
│   ├── Front_End/         # UI implementation
│   ├── Paper_Generator/   # Question generation logic
│   └── SQL_Database/      # Database operations
├── docs/                  # Documentation and test papers
└── requirements.txt       # Project dependencies
```

## Getting Started

### Prerequisites

- Python 3.8+
- MySQL Server
- Visual Studio Code or other IDE
- Anaconda or Miniconda (recommended for environment management)

### Installation

1. Clone this repository
```sh
git clone https://github.com/Angelocaduto927/JAP_LLM_Platform.git
cd JAP_LLM_Platform
```
2. Create a new conda environment:
```sh
conda create -n japllm python=3.10
```

3. Activate the environment:
```sh
conda activate japllm
```

4. Install dependencies
```sh
pip install -r requirements.txt
```

5. Configure database settings in `Source Code/SQL_Database/Excel2Db.py`:
```python
db = mysql.connector.connect(
    host="localhost",
    user="root", 
    password="your_password",
    database="japgpt"
)
```

6. Run the application
```sh
python Source Code/Front_End/interface.py
```

## Usage

1. Launch the application
2. Enter student information (name or ID)
3. Select test paper
4. Generate or grade papers
5. View results and feedback


## License

This project is licensed under the MIT License - see the LICENSE file for details.