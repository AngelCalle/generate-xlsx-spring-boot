# Generate .xlsx Spring Boot  


1. Once the project is executed, test data must be inserted into the database.  



#INSERT INTO USERSS VALUES (1, 'caramelo@caramelo.es', true, 'Caramelo', '1234');  

#INSERT INTO USERSS VALUES (2, 'piruleta@piruleta.es', true, 'Piruleta', '1234');  

#INSERT INTO ROLES VALUES (1, 'ADMIN', 'El Mandatario');  

#INSERT INTO ROLES VALUES (2, 'BASIC', 'El Mindundi');  

#INSERT INTO USERSS_ROLES VALUES (1, 1);  

#INSERT INTO USERSS_ROLES VALUES (1, 2);  

#INSERT INTO USERSS_ROLES VALUES (2, 2);  

#SELECT * FROM USERSS;  

#SELECT * FROM USERSS_ROLES;  

#SELECT * FROM USERSS_ROLES;  


2. Download an .xlsx report with the data from the database
    http://localhost:8080/report/download-excel

3. Save an .xlsx report to a specified path src\main\resources
    http://localhost:8080/report/report-excel
