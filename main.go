package main

import (
	"fmt"
	"log"
	"os"
	"sort"
	"strconv"

	"github.com/xuri/excelize/v2"
)

type StudentRecord struct {
	SlNo       int
	ClassNo    int
	Emplid     string
	CampusID   string
	Quiz       float64
	MidSem     float64
	LabTest    float64
	WeeklyLabs float64
	PreCompre  float64
	Compre     float64
	Total      float64
}

type BranchAverage struct {
	Branch string
	Total  float64
	Count  int
}

func main() {
	if len(os.Args) < 2 {
		log.Fatal("Please provide the path to the XLSX file as a command-line argument.")
	}

	filePath := os.Args[1]
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		log.Fatalf("Failed to open file: %s", err)
	}
	defer f.Close()

	rows, err := f.GetRows("CSF111_202425_01_GradeBook")
	if err != nil {
		log.Fatalf("Failed to get rows: %s", err)
	}

	var records []StudentRecord
	var discrepancies []string
	branchAverages := make(map[string]BranchAverage)

	// Branch Code Mapping
	branchMapping := map[string]string{
		"A7": "CS",
		"AA": "ECE",
		"A8": "ENI",
		"A3": "EEE",
		"A4": "MECH",
		"A5": "BPHARM",
		"AD": "MANU",
	}

	// Process each row
	for i, row := range rows {
		if i == 0 {
			continue // Skip header row
		}

		if len(row) < 11 {
			continue // Skip rows with insufficient data
		}

		record, err := parseRow(row)
		if err != nil {
			log.Printf("Error parsing row %d: %s", i+1, err)
			continue
		}

		// Check if the computed total matches the recorded total
		computedTotal := record.Quiz + record.MidSem + record.LabTest + record.WeeklyLabs + record.PreCompre + record.Compre
		if computedTotal != record.Total {
			discrepancies = append(discrepancies, fmt.Sprintf("Discrepancy for Emplid %s: Computed Total %.2f != Recorded Total %.2f", record.Emplid, computedTotal, record.Total))
		}

		records = append(records, record)

		// Process only students from 2024 batch
		if len(record.CampusID) >= 6 && record.CampusID[:4] == "2024" {
			branchCode := record.CampusID[4:6] // Extract branch code (e.g., "A7", "AA")

			if branchName, exists := branchMapping[branchCode]; exists {
				ba := branchAverages[branchName]
				ba.Branch = branchName
				ba.Total += record.Total
				ba.Count++
				branchAverages[branchName] = ba
			}
		}
	}

	// Calculate general averages
	var quizSum, midSemSum, labTestSum, weeklyLabsSum, preCompreSum, compreSum, totalSum float64
	for _, record := range records {
		quizSum += record.Quiz
		midSemSum += record.MidSem
		labTestSum += record.LabTest
		weeklyLabsSum += record.WeeklyLabs
		preCompreSum += record.PreCompre
		compreSum += record.Compre
		totalSum += record.Total
	}

	numRecords := float64(len(records))
	quizAvg := quizSum / numRecords
	midSemAvg := midSemSum / numRecords
	labTestAvg := labTestSum / numRecords
	weeklyLabsAvg := weeklyLabsSum / numRecords
	preCompreAvg := preCompreSum / numRecords
	compreAvg := compreSum / numRecords
	totalAvg := totalSum / numRecords

	// Print discrepancies
	if len(discrepancies) > 0 {
		fmt.Println("Discrepancies found:")
		for _, d := range discrepancies {
			fmt.Println(d)
		}
	} else {
		fmt.Println("No discrepancies found.")
	}

	// Print general averages
	fmt.Printf("\nGeneral Averages:\n")
	fmt.Printf("Quiz: %.2f\n", quizAvg)
	fmt.Printf("Mid-Sem: %.2f\n", midSemAvg)
	fmt.Printf("Lab Test: %.2f\n", labTestAvg)
	fmt.Printf("Weekly Labs: %.2f\n", weeklyLabsAvg)
	fmt.Printf("Pre-Compre: %.2f\n", preCompreAvg)
	fmt.Printf("Compre: %.2f\n", compreAvg)
	fmt.Printf("Total: %.2f\n", totalAvg)

	// Print branch-wise averages
	fmt.Printf("\nBranch-wise Averages (2024 Only):\n")
	for branch, ba := range branchAverages {
		avgTotal := ba.Total / float64(ba.Count)
		fmt.Printf("Branch average for %s is %.2f\n", branch, avgTotal)
	}

	// Print top 3 students for each component
	fmt.Printf("\nTop 3 Students:\n")
	printTopStudents(records, "Quiz", func(r StudentRecord) float64 { return r.Quiz })
	printTopStudents(records, "Mid-Sem", func(r StudentRecord) float64 { return r.MidSem })
	printTopStudents(records, "Lab Test", func(r StudentRecord) float64 { return r.LabTest })
	printTopStudents(records, "Weekly Labs", func(r StudentRecord) float64 { return r.WeeklyLabs })
	printTopStudents(records, "Pre-Compre", func(r StudentRecord) float64 { return r.PreCompre })
	printTopStudents(records, "Compre", func(r StudentRecord) float64 { return r.Compre })
	printTopStudents(records, "Total", func(r StudentRecord) float64 { return r.Total })
}

func parseRow(row []string) (StudentRecord, error) {
	var record StudentRecord
	var err error

	record.SlNo, err = strconv.Atoi(row[0])
	if err != nil {
		return record, fmt.Errorf("invalid Sl No: %s", row[0])
	}

	record.ClassNo, err = strconv.Atoi(row[1])
	if err != nil {
		return record, fmt.Errorf("invalid Class No: %s", row[1])
	}

	record.Emplid = row[2]
	record.CampusID = row[3]

	record.Quiz, err = strconv.ParseFloat(row[4], 64)
	if err != nil {
		return record, fmt.Errorf("invalid Quiz: %s", row[4])
	}

	record.MidSem, err = strconv.ParseFloat(row[5], 64)
	if err != nil {
		return record, fmt.Errorf("invalid Mid-Sem: %s", row[5])
	}

	record.LabTest, err = strconv.ParseFloat(row[6], 64)
	if err != nil {
		return record, fmt.Errorf("invalid Lab Test: %s", row[6])
	}

	record.WeeklyLabs, err = strconv.ParseFloat(row[7], 64)
	if err != nil {
		return record, fmt.Errorf("invalid Weekly Labs: %s", row[7])
	}

	record.PreCompre, err = strconv.ParseFloat(row[8], 64)
	if err != nil {
		return record, fmt.Errorf("invalid Pre-Compre: %s", row[8])
	}

	record.Compre, err = strconv.ParseFloat(row[9], 64)
	if err != nil {
		return record, fmt.Errorf("invalid Compre: %s", row[9])
	}

	record.Total, err = strconv.ParseFloat(row[10], 64)
	if err != nil {
		return record, fmt.Errorf("invalid Total: %s", row[10])
	}

	return record, nil
}

func printTopStudents(records []StudentRecord, component string, getScore func(StudentRecord) float64) {
	sort.Slice(records, func(i, j int) bool {
		return getScore(records[i]) > getScore(records[j])
	})

	fmt.Printf("\nTop 3 Students for %s:\n", component)
	for i := 0; i < 3 && i < len(records); i++ {
		fmt.Printf("%d. Emplid: %s, Marks: %.2f\n", i+1, records[i].Emplid, getScore(records[i]))
	}
}
