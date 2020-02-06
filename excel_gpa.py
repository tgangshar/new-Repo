import xlwt
from xlwt import Workbook
import pathlib

# Workbook is created
wb = Workbook()

# add_sheet is used to create sheet.
student = wb.add_sheet('Tenzin')


def notCore(clas):
	nonCore = ("ECON, MATH")
	if clas in nonCore:
		return True
	return False


def isAlsoCore(clas):
	alsoCore = ("PHIL1000", "PHIL3000")
	if clas in alsoCore:
		return True
	return False


def get_length(stdnt_clss):
	length = 0
	for clas in stdnt_clss:
		if len(clas) > 0:
			length += 1
	return length


def write_all(Student_Classes, Student_Grades, Student_Credit):
	# row = "{:12}{:14}{:10}{:10}{:10}"
	# print(row.format("Course", 'Year', 'Grade', 'Credits', 'Requirement'))
	number_of_classes = get_length(Student_Classes)

	table = 2
	major = "CISC"
	minor = "PHIL"
	master = "CISC"
	maj_count = 0
	core_count = 0
	minor_count = 0
	for i in range(0, number_of_classes):
		course = Student_Classes[i]
		if len(course) > 8:
			curr_year = course
			course = curr_year[:8]
			curr_year = curr_year[9:]

		if notCore(course):
			req = "Nothing"
		elif major in course:
			maj_count += 1
			req = 'Major'
		elif minor in course:
			minor_count += 1
			req = minor
			if isAlsoCore(course):
				req += " & Core"
		else:
			core_count += 1
			req = 'Core'
		# print(row.format(course, curr_year, Student_Grades[i], Student_Credits[i], req))
		student.write(table, 0, course)
		student.write(table, 1, curr_year)
		student.write(table, 2, Student_Grades[i])
		student.write(table, 3, int(Student_Credit[i]))
		student.write(table, 4, req)
		table += 1


# print("MAJOR:", maj_count, "CORE:", core_count, "MINOR:", minor_count)


def calc_cmu(student):
	gpa = {'A': 4.0, 'A-': 3.7, 'B+': 3.33, 'B': 3.0, 'B-': 2.7,
	       'C+': 2.3, 'C': 2.3, 'C-': 2.0, 'D+': 1.7, 'D': 1.0, 'D-': 0.7}
	GradePoint = 0
	credits_completed = 6
	for grde, cred in zip(student.col[2], student.col[3]):
		creds = int(cred)
		if isinstance(grde, str):
			grde = gpa[grde]
		GradePoint = GradePoint + grde * creds
		credits_completed += creds
	return GradePoint / credits_completed


student.write(1, 0, "Course")
student.write(1, 1, 'Year')
student.write(1, 2, 'Grade')
student.write(1, 3, 'Credits')
student.write(1, 4, 'Requirement')

# inserting Grades based on semesters

credits_completed = 6
cumulative_gpa = 0

maj_needed = 15
minor_needed = 6
core_needed = 13

Student_Classes = ("CISC1600 Fall2018", "CISC1610", "ECON1100", "ENGL1102", "HIST1300", "MATH1206",
                   "CISC2000 Spring2019", "CISC2010", "MATH2001", "PHIL1000", "THEO1000",
                   "CISC2200 Fall2019", "CISC2500", "NSCI1501", "NSCI1502", "PHIL3000", "THEO3724")

Student_Grades = ("B-", "B-", "D", "B-", "C", "B",
                  "A-", "B+", "B-", "B+", "A-",
                  "B+", "A-", "B+", "A-", "A", "A-")
Student_Credits = ("3", "1", "3", "3", "3", "4",
                   "3", "1", "4", "3", "3",
                   "4", "4", "3", "1", "3", "3")

write_all(Student_Classes, Student_Grades, Student_Credits)

print("Cumulative GPA:", calc_cmu(student))

wb.save("curr_Course.xls")
current_path = pathlib.Path().absolute()
print(current_path)
