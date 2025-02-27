package main

import (
	"fmt"
	"log"
	"math"
	"os"
	"sort"
	"strconv"

	"github.com/xuri/excelize/v2"
)

// Student structure
type Student struct {
	EmpID      string
	Branch     string
	Quiz       float64
	MidSem     float64
	LabTest    float64
	WeeklyLabs float64
	Compre     float64
	Total      float64
}

// Branch name mapping
var branchMap = map[string]string{
	"2021A2": "Civil 2021", "2024A3": "EEE 2024", "2024A4": "Mechanical 2024",
	"2024A5": "Pharma 2024", "2024A7": "CSE 2024", "2024A8": "ENI 2024", "2024AA": "ECE 2024",
	"2024AD": "MnC 2024", "2024B1": "MSc Biology", "2020B5": "MSc Physics 2020", "2021A7": "CSE 2021", "2022A7": "CSE 2022",
	"2023A7": "CSE 2023", "2021A8": "ENI 2021", "2021AA": "ECE 2021", "2021B1": "Msc Biology 2021", "2021B4": "Msc Maths 2021",
	"2021B5": "Msc Physics 2021", "2022A1": "Chemical 2022", "2022A2": "Civil 2022", "2022A3": "EEE 2022", "2022A4": "Mechanical 2022",
	"2022AA": "ECE 2022", "2022B2": "MSc Chemistry 2022", "2023A5": "Pharma 2023", "2023A8": "ENI 2023",
}

const tolerance = 0.01 // handling floating point precision

func main() {
	if len(os.Args) < 2 {
		fmt.Println("Usage - go run main.go <path-to-file.xlsx>")
		os.Exit(1)
	}

	filePath := os.Args[1]

	students, branchSums, branchCounts, totalSum, totalCount := processFile(filePath)

	printResults(students, branchSums, branchCounts, totalSum, totalCount)
}

// Processes the Excel file and returns the necessary data
func processFile(filePath string) ([]Student, map[string]float64, map[string]int, float64, int) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		log.Fatalf("Failed to open file: %v", err)
	}
	defer f.Close()

	sheetName := f.GetSheetName(0)
	rows, err := f.GetRows(sheetName)
	if err != nil {
		log.Fatalf("Failed to read rows: %v", err)
	}

	var students []Student
	branchSums := make(map[string]float64)
	branchCounts := make(map[string]int)
	var totalSum float64
	var totalCount int

	for i, row := range rows {
		if i == 0 || len(row) < 10 {
			continue
		}

		student, valid := parseRow(row)
		if !valid {
			continue
		}

		students = append(students, student)
		branchSums[student.Branch] += student.Total
		branchCounts[student.Branch]++
		totalSum += student.Total
		totalCount++
	}

	return students, branchSums, branchCounts, totalSum, totalCount
}

// Parses a row from the Excel file and returns a Student struct and a validity flag
func parseRow(row []string) (Student, bool) {
	empID := row[2]
	campusID := row[3]
	quiz, _ := strconv.ParseFloat(row[4], 64)
	midSem, _ := strconv.ParseFloat(row[5], 64)
	labTest, _ := strconv.ParseFloat(row[6], 64)
	weeklyLabs, _ := strconv.ParseFloat(row[7], 64)
	compre, _ := strconv.ParseFloat(row[9], 64)
	total, _ := strconv.ParseFloat(row[10], 64)

	branch := extractBranch(campusID)
	if len(branch) < 6 {
		log.Printf("Skipping row due to invalid branch ID: %s\n", campusID)
		return Student{}, false
	}

	preCompre := quiz + midSem + labTest + weeklyLabs
	calculatedTotal := preCompre + compre

	if !isWithinTolerance(calculatedTotal, total) {
		log.Printf("Discrepancy in total marks for EmpID %s: Expected %.2f, Found %.2f\n",
			empID, calculatedTotal, total)
	}

	student := Student{
		EmpID:      empID,
		Branch:     branch,
		Quiz:       quiz,
		MidSem:     midSem,
		LabTest:    labTest,
		WeeklyLabs: weeklyLabs,
		Compre:     compre,
		Total:      total,
	}

	return student, true
}

// Extracts branch from Campus ID
func extractBranch(campusID string) string {
	if len(campusID) < 6 {
		return ""
	}
	branch := campusID[:6]
	if _, exists := branchMap[branch]; exists {
		return branch
	}
	return ""
}

// Checks if two floating-point numbers are within a specified tolerance
func isWithinTolerance(a, b float64) bool {
	return math.Abs(a-b) <= tolerance
}

// Prints the results
func printResults(students []Student, branchSums map[string]float64, branchCounts map[string]int, totalSum float64, totalCount int) {
	fmt.Println("======================================")
	fmt.Println("Top 3 Students for Each Component")
	printTopStudents(students)

	fmt.Println("\n======================================")
	fmt.Println("Overall and Branch-Wise Averages")
	fmt.Printf("Overall Average Marks: %.2f\n", totalSum/float64(totalCount))
	for branch, sum := range branchSums {
		fmt.Printf("Branch %s (%s) Average Marks: %.2f\n", branch, branchMap[branch], sum/float64(branchCounts[branch]))
	}
}

// Prints top 3 students for each component
func printTopStudents(students []Student) {
	components := []struct {
		name   string
		getVal func(Student) float64
	}{
		{"Quiz (30)", func(s Student) float64 { return s.Quiz }},
		{"Mid-Sem (75)", func(s Student) float64 { return s.MidSem }},
		{"Lab Test (60)", func(s Student) float64 { return s.LabTest }},
		{"Weekly Labs", func(s Student) float64 { return s.WeeklyLabs }},
		{"Compre (105)", func(s Student) float64 { return s.Compre }},
		{"Total (300)", func(s Student) float64 { return s.Total }},
	}

	for _, comp := range components {
		fmt.Printf("\nTop 3 for %s:\n", comp.name)
		sorted := sortByComponent(students, comp.getVal)
		for i, s := range sorted[:min(3, len(sorted))] {
			fmt.Printf("%d. EmpID: %s - %.2f\n", i+1, s.EmpID, comp.getVal(s))
		}
	}
}

// Sorts students by a given component using sort.Slice
func sortByComponent(students []Student, getVal func(Student) float64) []Student {
	sorted := append([]Student{}, students...)
	sort.Slice(sorted, func(i, j int) bool {
		return getVal(sorted[i]) > getVal(sorted[j])
	})
	return sorted
}

// Returns the minimum of two numbers
func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}
