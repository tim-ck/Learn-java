# little side project for me to study apache poi

This project is a Java standalone application that focuses on studying object-oriented programming (OOP) and Excel manipulation using poi apache.

## Project Description

The main objective of this project is to handle a collection of records and export them to Excel files. The challenge lies in the fact that different types of records need to be outputted to separate Excel files. However, these different record types share common attributes, which are stored in a common file.

## Features

- Importing and processing records
- Generating Excel files based on record types
- Create template file with drop down cell and using a reference sheet to store validation data
- import data to template file
- Protect sheet


## Build Instructions

To build the program, follow these steps:

1. Make sure you have Maven installed on your system.
2. Clone the repository
3. Run the following command to clean and install the project:
    
```bash
mvn clean install
```
5. The build will create a JAR file in the `target` directory.

## Run Instructions

To run the program, run the following command to execute the JAR file:

```bash
    java -jar my-app-jar-with-dependencies.jar
```



## Usage

1. Modify the common file to define the shared attributes for all record types.
2. Implement specific record types by extending the common file and adding additional attributes.
3. Use the provided methods to import and process records.
4. Call the appropriate methods to generate Excel files based on record types.
