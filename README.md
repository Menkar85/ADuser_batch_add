# Batch User Creation in Active Directory (Work in progress)

[![GitHub License](https://img.shields.io/github/license/menkar85/ADuser_batch_add)](LICENSE)
[![GitHub issues](https://img.shields.io/github/issues/menkar85/ADuser_batch_add)](https://github.com/menkar85/ADuser_batch_add/issues)
[![GitHub stars](https://img.shields.io/github/stars/menkar85/ADuser_batch_add)](https://github.com/menkar85/ADuser_batch_add/stargazers)

A tool to automate the creation of multiple user accounts in Active Directory. This application fetch data from Excel file and creates users in Active Directory domain.

## Features

- **Batch Processing**: Create multiple user accounts at once.
- **XLSX Import**: Read user details from a XLSX template file.
- **Customizable Attributes**: Configure attributes like Last Name, Full name.
- **Error Handling**: Provides error logs for troubleshooting.

## Prerequisites

Before using the tool, ensure you have:

- Access to an Active Directory environment.
- Administrative privileges to create new user accounts.
- Python installed on your system (version 3.x recommended).

## Installation

1. Clone the repository:
    ```ps1
    git clone https://github.com/menkar85/ADuser_batch_add.git
    cd ADuser_batch_add
    ```

2. Install dependencies:
    ```ps1
    pip install -r requirements.txt
    ```

3. Start the application and add required information.

## Usage

1. Make a copy of a template file, adjust up to your needs and save.
2. Start the app, insert required info and press Start Import
3. Check result file and log file for errors.

Destination ou should be in folder format where "/" is root domain/subdomain root. 
User container will be recorded as User, Staff container inside User container will be
recorded as User/Staff and so on. 

In case there is no such container - it will be created. It is highly recommended to created users into OU which first 
to be created as the app will raise an error in case of username duplicates. In case of importing users into 
existing container all duplicated will be marked as existing and no error will be raised.

If you got duplicates error during import it is recommended to resolve all issues in source xlsx file, delete destination 
OU and reimport again in order to get notified about duplicated. Such behavior is resulting from idempotent nature of 
the app.