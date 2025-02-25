package main

import (
	"fmt"
	"log"
	"math"
	"strconv"

	"github.com/xuri/excelize/v2"
)

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

var branchMap = map[string]string{
	"A1": "Chemical", "A2": "Civil", "A3": "EEE", "A4": "Mechanical",
	"A5": "Pharma", "A7": "CSE", "A8": "ENI", "AA": "ECE",
	"AB": "Manufacturing", "AD": "MnC", "B1": "MSc Biology", "B2": "MSc Chemistry",
	"B3": "MSc Economics", "B4": "MScMaths", "B5": "MSc Physics",
}

const tolerance = 0.01 //handling floating point precision

func main() {
	filePath := "C:\\Users\\Ditya\\Downloads\\PostmanBackendTask1_Data.xlsx"

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

		empID := row[2]
		campusID := row[3]
		quiz, _ := strconv.ParseFloat(row[4], 64)
		midSem, _ := strconv.ParseFloat(row[5], 64)
		labTest, _ := strconv.ParseFloat(row[6], 64)
		weeklyLabs, _ := strconv.ParseFloat(row[7], 64)
		compre, _ := strconv.ParseFloat(row[9], 64)
		total, _ := strconv.ParseFloat(row[10], 64)

		branch := extractBranch(campusID)
		if branch == "" {
			log.Printf("Skipping row %d due to invalid branch ID: %s\n", i+1, campusID)
			continue
		}

		// Calculate Pre-Compre and validate
		preCompre := quiz + midSem + labTest + weeklyLabs
		calculatedTotal := preCompre + compre

		if math.Abs(calculatedTotal-total) > tolerance {
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
		students = append(students, student)

		branchSums[branch] += total
		branchCounts[branch]++
		totalSum += total
		totalCount++
	}

	fmt.Println("======================================")
	fmt.Println(" Top 3 Students for Each Component ")
	printTopStudents(students)

	fmt.Println("\n======================================")
	fmt.Println(" Overall and Branch-Wise Averages ")
	fmt.Printf("Overall Average Marks: %.2f\n", totalSum/float64(totalCount))
	for branch, sum := range branchSums {
		fmt.Printf("Branch %s (%s) Average Marks: %.2f\n", branch, branchMap[branch], sum/float64(branchCounts[branch]))
	}
}

// Extracts branch from Campus ID
func extractBranch(campusID string) string {
	if len(campusID) < 6 {
		return ""
	}
	branch := campusID[4:6]
	if _, exists := branchMap[branch]; exists {
		return branch
	}
	return ""
}

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

func sortByComponent(students []Student, getVal func(Student) float64) []Student {
	sorted := append([]Student{}, students...)
	for i := 0; i < len(sorted)-1; i++ {
		for j := 0; j < len(sorted)-i-1; j++ {
			if getVal(sorted[j]) < getVal(sorted[j+1]) {
				sorted[j], sorted[j+1] = sorted[j+1], sorted[j]
			}
		}
	}
	return sorted
}

func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}
