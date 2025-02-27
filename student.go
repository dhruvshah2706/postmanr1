package main

import (
	// "errors"
	"fmt"
	"os"
	"sort"
	"strconv"
	"math"
	"strings"
	"github.com/xuri/excelize/v2"
)

type Student struct {
	SlNo        int
	ClassNo     int
	Emplid      string
	CampusID    string
	Quiz        float64
	MidSem      float64
	LabTest     float64
	WeeklyLabs  float64
	PreCompre   float64
	Compre      float64
	Total       float64
	ComputedSum float64
}

// Parses the Excel file and extracts student records
func parseExcel(filePath string) ([]Student, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open file: %w", err)
	}
	defer f.Close()

	rows, err := f.GetRows(f.GetSheetName(0))
	if err != nil {
		return nil, fmt.Errorf("failed to read sheet: %w", err)
	}

	var students []Student

	// Read rows (Skip header row)
	for i, row := range rows[1:] {
		if len(row) < 11 { // Ensure the row has enough columns
			continue
		}

		student, err := parseRow(i+2, row) // i+2 to map to actual row number in Excel
		if err != nil {
			fmt.Printf("Error parsing row %d: %v\n", i+2, err)
			continue
		}

		students = append(students, student)
	}

	return students, nil
}

func almostEqual(a, b float64) bool {
	const epsilon = 0.01 // Allow small difference due to precision
	return math.Abs(a-b) < epsilon
}

// Parses a row into a Student struct
func parseRow(rowNum int, row []string) (Student, error) {
	parseFloat := func(s string) (float64, error) {
		if s == "" {
			return 0, nil
		}
		return strconv.ParseFloat(s, 64)
	}

	slNo, err := strconv.Atoi(row[0])
	if err != nil {
		return Student{}, fmt.Errorf("invalid Sl No at row %d", rowNum)
	}
	classNo, err := strconv.Atoi(row[1])
	if err != nil {
		return Student{}, fmt.Errorf("invalid Class No at row %d", rowNum)
	}

	quiz, err := parseFloat(row[4])
	midSem, err := parseFloat(row[5])
	labTest, err := parseFloat(row[6])
	weeklyLabs, err := parseFloat(row[7])
	preCompre, err := parseFloat(row[8])
	compre, err := parseFloat(row[9])
	total, err := parseFloat(row[10])

	if err != nil {
		return Student{}, fmt.Errorf("invalid numeric data at row %d", rowNum)
	}
	

	

	// Validate PreCompre sum
	computedPreCompre := quiz + midSem + labTest + weeklyLabs
	if !almostEqual(computedPreCompre, preCompre) {
		fmt.Printf("Error: Mismatch in PreCompre at row %d. Expected %.2f, Found %.2f\n", rowNum, computedPreCompre, preCompre)
	}

	// Validate total sum
	computedSum := quiz + midSem + labTest + weeklyLabs + compre
	if !almostEqual(computedSum, total) {
		fmt.Printf("Error: Mismatch in total at row %d. Expected %.2f, Found %.2f\n", rowNum, computedSum, total)
	}
	return Student{
		SlNo:        slNo,
		ClassNo:     classNo,
		Emplid:      row[2],
		CampusID:    row[3],
		Quiz:        quiz,
		MidSem:      midSem,
		LabTest:     labTest,
		WeeklyLabs:  weeklyLabs,
		PreCompre:   preCompre,
		Compre:      compre,
		Total:       total,
		ComputedSum: computedSum,
	}, nil
}

// Computes averages for each component
func computeAverages(students []Student) map[string]float64 {
	sum := make(map[string]float64)
	count := float64(len(students))

	for _, s := range students {
		sum["Quiz"] += s.Quiz
		sum["MidSem"] += s.MidSem
		sum["LabTest"] += s.LabTest
		sum["WeeklyLabs"] += s.WeeklyLabs
		sum["PreCompre"] += s.PreCompre
		sum["Compre"] += s.Compre
		sum["Total"] += s.Total
	}

	averages := make(map[string]float64)
	for key, val := range sum {
		averages[key] = val / count
	}

	return averages
}

// Computes branch-wise average (based on 2024 batch)
func computeBranchAverages(students []Student) map[string]float64 {
	branchTotal := make(map[string]float64)
	branchCount := make(map[string]int)

	for _, s := range students {
		if len(s.CampusID) < 6 {
			continue // Skip invalid CampusID
		}
		yearPrefix := s.CampusID[:4]   // "2024"
		branchCode := s.CampusID[4:6] 

		if yearPrefix == "2024" && strings.Contains(branchCode, "A") {
			branchTotal[branchCode] += s.Total
			branchCount[branchCode]++
		}
	}

	// Compute average per branch
	branchAvg := make(map[string]float64)
	for branch, total := range branchTotal {
		branchAvg[branch] = total / float64(branchCount[branch])
	}

	return branchAvg
}


// Determines top 3 students for each component
func rankStudents(students []Student) map[string][]Student {
	categories := []string{"Quiz", "MidSem", "LabTest", "WeeklyLabs", "PreCompre", "Compre", "Total"}
	rankings := make(map[string][]Student)

	for _, category := range categories {
		// Create a copy of students to avoid modifying the original order
		sortedStudents := append([]Student{}, students...) 

		sort.Slice(sortedStudents, func(i, j int) bool {
			return getScoreByCategory(sortedStudents[i], category) > getScoreByCategory(sortedStudents[j], category)
		})
		rankings[category] = sortedStudents[:min(3, len(sortedStudents))]
	}

	return rankings
}

// Retrieves score for a specific category
func getScoreByCategory(student Student, category string) float64 {
	switch category {
	case "Quiz":
		return student.Quiz
	case "MidSem":
		return student.MidSem
	case "LabTest":
		return student.LabTest
	case "WeeklyLabs":
		return student.WeeklyLabs
	case "PreCompre":
		return student.PreCompre
	case "Compre":
		return student.Compre
	case "Total":
		return student.Total
	}
	return 0
}

// Utility function to get minimum value
func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}

func main() {
	if len(os.Args) < 2 {
		fmt.Println("Usage: go run main.go <path-to-excel-file>")
		return
	}

	filePath := os.Args[1]
	students, err := parseExcel(filePath)
	if err != nil {
		fmt.Println("Error:", err)
		return
	}

	fmt.Println("\n--- Average Scores ---")
	averages := computeAverages(students)
	for key, val := range averages {
		fmt.Printf("%s: %.2f\n", key, val)
	}

	fmt.Println("\n--- Branch-wise Averages (2024 Batch) ---")
	branchAverages := computeBranchAverages(students)
	for  branch, avg := range branchAverages {
		fmt.Printf("Branch %s: %.2f\n", branch, avg)
	}

	fmt.Println("\n--- Top 3 Students Per Component ---")
	rankings := rankStudents(students)
	for category, topStudents := range rankings {
		fmt.Printf("\n%s:\n", category)
		for rank, student := range topStudents {
			fmt.Printf("Rank:%d. %s - %.2f\n", rank+1, student.Emplid, getScoreByCategory(student, category))
		}
	}
}
