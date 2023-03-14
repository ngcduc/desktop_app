CREATE DATABASE hospital_managemnet_system

USE hospital_managemnet_system

CREATE TABLE hospital_managemnet_system.`hospital` (
  `id` int NOT NULL AUTO_INCREMENT,
  `name` varchar(500) NOT NULL,
  `phone` varchar(10) DEFAULT NULL,
  `address` varchar(200) DEFAULT NULL,
  `description` varchar(500) DEFAULT NULL,
  PRIMARY KEY (`id`)
) 

CREATE TABLE hospital_managemnet_system.`doctor` (
  `id` int NOT NULL AUTO_INCREMENT,
  `name` varchar(500) NOT NULL,
  `phone` varchar(10) DEFAULT NULL,
  `email` varchar(100) NOT NULL,
  `address` varchar(300) DEFAULT NULL,
  `hospital_id` int NOT NULL,
  PRIMARY KEY (`id`),
  KEY `doctor_FK` (`hospital_id`),
  CONSTRAINT `doctor_FK` FOREIGN KEY (`hospital_id`) REFERENCES `hospital` (`id`) ON DELETE CASCADE ON UPDATE CASCADE
)


CREATE TABLE hospital_managemnet_system.`patient` (
  `id` int NOT NULL AUTO_INCREMENT,
  `name` varchar(200) NOT NULL,
  `phone` varchar(10) DEFAULT NULL,
  `email` varchar(200) NOT NULL,
  `address` varchar(200) NOT NULL,
  `hospital_id` int NOT NULL,
  PRIMARY KEY (`id`),
  KEY `patient_FK` (`hospital_id`),
  CONSTRAINT `patient_FK` FOREIGN KEY (`hospital_id`) REFERENCES `hospital` (`id`) ON DELETE CASCADE ON UPDATE CASCADE
)

CREATE TABLE hospital_managemnet_system.`schedule` (
  `id` int NOT NULL AUTO_INCREMENT,
  `name` varchar(200) NOT NULL,
  `date` datetime NOT NULL,
  `doctor_id` int NOT NULL,
  `patient_id` int NOT NULL,
  `result` varchar(300) DEFAULT NULL,
  PRIMARY KEY (`id`),
  KEY `schedule_FK` (`patient_id`),
  KEY `schedule_FK_1` (`doctor_id`),
  CONSTRAINT `schedule_FK` FOREIGN KEY (`patient_id`) REFERENCES `patient` (`id`) ON DELETE CASCADE ON UPDATE CASCADE,
  CONSTRAINT `schedule_FK_1` FOREIGN KEY (`doctor_id`) REFERENCES `doctor` (`id`) ON DELETE CASCADE ON UPDATE CASCADE
)