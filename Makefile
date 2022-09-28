PYTHON = python3

.PHONY = help test test-coverage

.DEFAULT_GOAL = help

help:
	@echo "-----------------HELP-----------------"
	@echo "make test          : run all project tests"
	@echo "make test-file     : run test files specified by filling in FILE="""
	@echo "                     with file names"
	@echo "make test-coverage : run the coverage tests"

	@echo "--------------------------------------"

test:
	${PYTHON} -m unittest

test-file: 
	for file in ${FILES} ; do \
		${PYTHON} -m unittest tests/$$file.py ; \
	done

test-performance:
	${PYTHON} -m unittest tests/stress_tests.py

test-coverage:
	coverage run -m unittest
	coverage report