#!/bin/bash

rm export/*.xlsx
mvn clean compile exec:java
