# Research Assistant Opportunity Application Automation

Welcome to the Research Assistant (RA) Opportunity Application Automation project! This program is designed to simplify the process of applying for research assistant positions by automating the collection and submission of information from your university's faculty directory to a Google Form.

## Table of Contents
- [Project Overview](#project-overview)
- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
- [Contributing](#contributing)
- [License](#license)

## Project Overview
Are you a student seeking research assistant opportunities? Tired of manually inputting your details in countless application forms? This project aims to save you time and effort by streamlining the process. By utilizing web scraping techniques, the program extracts the necessary information from your university's faculty directory and fills out a Google Form automatically. Additionally, it sends an email expressing your interest in the research assistant position to the respective faculty members.

## Features
- Web scraping of university's faculty directory.
- Automatic form filling using Google Forms.
- Customizable email template for expressing interest.
- User-friendly configuration and input management.

## Installation
1. Clone the repository using the following command:
```bash
   git clone https://github.com/guccigamp/RA-Apply.git
```
2. Install the required dependencies:    
```bash
    pip install -r requirements.txt
```
## Usage
1. Configure the project settings in the config.json file. Input your Google Form URL, email details, and other preferences.
2. Run the program:
 ```bash
    python scraper.py
```
3. Once done with the execution, convert the form responses into a google sheet and download it.
4. Create an email template in txt file. You can personalise it by adding "[Professor's Last Name]" fields in the template. 
 ### Note: It is possible to send dynamic emails but for the simplicity of the project we send text-only emails. 
5. Run the program and follow the given instructions: 
```bash
python main.py
```
## Contributing
Contributions are welcome and appreciated! If you have any improvements or feature suggestions, feel free to open an issue or submit a pull request. Please review the Contributing Guidelines before getting started.

## License
This project is licensed under the MIT License.