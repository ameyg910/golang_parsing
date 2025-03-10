/*
My explanation for the code in few sentences: 
1) Iterating through rows, checking for incomplete entries and the top header and skipping them
2)Row values -> Appropriate data types and storing them in the Student_Record structure
3)Checking the Total_Score by calculating and checking for discrepancies and then removing them
4) Creating Average variables for every test format and then computing averages for them
5) Calculating Branch wise average was tricky, I analyzed a way to check for branch codes and then mapping them to respective branches, I immediately thought of map data structure to perform this task.
6) To calculate top 3 rankers were easy, I thought of sorting the scores of every test format, and got to know about the sort lib in go
*/
package main

import (
	"fmt"
	"log"
	"os"
	"sort"
	"strconv"
	"github.com/xuri/excelize/v2"
)

type Student_Record struct { 
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
	Total_Score      float64
}

type BranchAverage struct {
	Branch string
	Total_Score  float64
	totalnoofstudents  int
}

func main() {
	if len(os.Args) < 2 {
		log.Fatal("Please provide the path to the XLSX file as a command-line argument.")
	}

	filePath := os.Args[1] 
	f, err := excelize.OpenFile(filePath) // Opening the file for the task to be executed
	if err != nil {
		log.Fatalf("Failed to open file: %s", err)
	}
	defer f.Close() // closing the file after the required operations are executed.

	rows, err := f.GetRows("CSF111_202425_01_GradeBook") // Reads all Rows from the given xlsx file
	if err != nil { // if not then it shall throw the given error
		log.Fatalf("Failed to get rows: %s", err)
	}

	var records []Student_Record // dynamic array of the S_R struct type (initializes to nil) created to store student records
	var discrepancies []string // DA to store all the inconsistencies in the total scores
	branchAverages := make(map[string]BranchAverage) // Stores branch-wise total and number of students

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

	for i, row := range rows {
		if i == 0 {
			continue // The top Header row shall be skipped 
		}

		if len(row) < 11 { // Skipping rows with insufficient data (11 cols should be there)
			continue 
		}

		record, err := parseRow(row) // Row -> Student_Record object 
		if err != nil {
			log.Printf("error parsing row %d: %s", i+1, err)
			continue
		}

		computedTotal_Score := record.Quiz + record.MidSem + record.LabTest + record.WeeklyLabs + record.PreCompre + record.Compre
		if computedTotal_Score != record.Total_Score { // if sum of individual scores != recorded total, add it to discrepancies 
			discrepancies = append(discrepancies, fmt.Sprintf("Discrepancy for Emplid %s: Computed Total_Score %.2f != Recorded Total_Score %.2f", record.Emplid, computedTotal_Score, record.Total_Score))
		}

		records = append(records, record)

		if len(record.CampusID) >= 6 && record.CampusID[:4] == "2024" {
			branchCode := record.CampusID[4:6] // Branch code extraction

			if branchName, exists := branchMapping[branchCode]; exists {
				ba := branchAverages[branchName]
				ba.Branch = branchName
				ba.Total_Score += record.Total_Score
				ba.totalnoofstudents++
				branchAverages[branchName] = ba
			}// iterating, Updating the branch's total scores and no.of students 
		}
	}

	var quizSum, midSemSum, labTestSum, weeklyLabsSum, preCompreSum, compreSum, Total_ScoreSum float64
	for _, record := range records {
		quizSum += record.Quiz
		midSemSum += record.MidSem
		labTestSum += record.LabTest
		weeklyLabsSum += record.WeeklyLabs
		preCompreSum += record.PreCompre
		compreSum += record.Compre
		Total_ScoreSum += record.Total_Score
	}// sum of all the test formats 

	numRecords := float64(len(records))
	quizAvg := quizSum / numRecords
	midSemAvg := midSemSum / numRecords
	labTestAvg := labTestSum / numRecords
	weeklyLabsAvg := weeklyLabsSum / numRecords
	preCompreAvg := preCompreSum / numRecords
	compreAvg := compreSum / numRecords
	Total_ScoreAvg := Total_ScoreSum / numRecords // average computation 

	if len(discrepancies) > 0 {
		fmt.Println("Discrepancies found:")
		for _, d := range discrepancies {
			fmt.Println(d)
		}
	} else {
		fmt.Println("No discrepancies found.")
	}

	fmt.Printf("\nGeneral Averages:\n")
	fmt.Printf("Quiz: %.2f\n", quizAvg)
	fmt.Printf("Mid-Sem: %.2f\n", midSemAvg)
	fmt.Printf("Lab Test: %.2f\n", labTestAvg)
	fmt.Printf("Weekly Labs: %.2f\n", weeklyLabsAvg)
	fmt.Printf("Pre-Compre: %.2f\n", preCompreAvg)
	fmt.Printf("Compre: %.2f\n", compreAvg)
	fmt.Printf("Total_Score: %.2f\n", Total_ScoreAvg)

	fmt.Printf("\nBranch-wise Averages (2024 Only):\n")
	for branch, ba := range branchAverages {
		avgTotal_Score := ba.Total_Score / float64(ba.totalnoofstudents)
		fmt.Printf("Branch average for %s is %.2f\n", branch, avgTotal_Score)
	}

	fmt.Printf("\nTop 3 Students:\n")
	printTopStudents(records, "Quiz", func(r Student_Record) float64 { return r.Quiz })
	printTopStudents(records, "Mid-Sem", func(r Student_Record) float64 { return r.MidSem })
	printTopStudents(records, "Lab Test", func(r Student_Record) float64 { return r.LabTest })
	printTopStudents(records, "Weekly Labs", func(r Student_Record) float64 { return r.WeeklyLabs })
	printTopStudents(records, "Pre-Compre", func(r Student_Record) float64 { return r.PreCompre })
	printTopStudents(records, "Compre", func(r Student_Record) float64 { return r.Compre })
	printTopStudents(records, "Total_Score", func(r Student_Record) float64 { return r.Total_Score })
}

func parseRow(row []string) (Student_Record, error) {
	var record Student_Record
	var err error
    // conversion of ASCII to integer and checking for invalid items (if any)
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

	record.Total_Score, err = strconv.ParseFloat(row[10], 64)
	if err != nil {
		return record, fmt.Errorf("invalid Total_Score: %s", row[10])
	}

	return record, nil
}

func printTopStudents(records []Student_Record, component string, getScore func(Student_Record) float64) {
	sort.Slice(records, func(i, j int) bool {
		return getScore(records[i]) > getScore(records[j])
	})

	fmt.Printf("\nTop 3 Students for %s:\n", component)
	for i := 0; i < 3 && i < len(records); i++ {
		fmt.Printf("%d. Emplid: %s, Marks: %.2f\n", i+1, records[i].Emplid, getScore(records[i]))
	}
} 
