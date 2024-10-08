import streamlit as st
import fitz  # PyMuPDF for reading PDFs
import re
from collections import defaultdict
import matplotlib.pyplot as plt
import pandas as pd  # For handling Excel files

# Function to extract text from PDF using PyMuPDF
def extract_text_from_pdf(pdf_file):
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")  # Read from BytesIO object
    full_text = ""
    for page_num in range(doc.page_count):
        page = doc.load_page(page_num)
        full_text += page.get_text("text")
    return full_text

# Function to extract questions, marks, and COs
def extract_question_data(text):
    question_pattern = re.compile(r"Q\.\d+")
    marks_pattern = re.compile(r"\[(\d+(\.\d+)?)\]")
    co_pattern = re.compile(r"CO\d+")

    questions = re.split(question_pattern, text)[1:]  # Splitting text by 'Q.'
    question_data = []
    
    for index, question in enumerate(questions):
        cos = co_pattern.findall(question)
        marks = marks_pattern.findall(question)
        
        total_marks = sum(float(mark[0]) for mark in marks)
        
        question_data.append({
            'question_number': index + 1,
            'cos': cos,
            'marks': total_marks
        })
    
    return question_data

# Function to aggregate marks by CO
def aggregate_marks_by_co(question_data):
    co_marks = defaultdict(float)

    for qdata in question_data:
        cos = qdata['cos']
        marks = qdata['marks']
        for co in cos:
            co_marks[co] += marks
    
    return co_marks

# Function to handle student data input manually
def get_student_data_manually():
    st.subheader("Enter Student Data Manually")
    student_data = []
    
    num_students = st.number_input("Number of Students", min_value=1, step=1)
    
    for i in range(num_students):
        with st.form(key=f'student_form_{i}'):
            enrollment_no = st.text_input(f"Enrollment No. {i + 1}")
            first_name = st.text_input(f"First Name {i + 1}")
            q1 = st.number_input(f"Q.1 Marks {i + 1}", min_value=0.0, format="%.2f")
            q2 = st.number_input(f"Q.2 Marks {i + 1}", min_value=0.0, format="%.2f")
            q3 = st.number_input(f"Q.3 Marks {i + 1}", min_value=0.0, format="%.2f")
            q4 = st.number_input(f"Q.4 Marks {i + 1}", min_value=0.0, format="%.2f")
            q5 = st.number_input(f"Q.5 Marks {i + 1}", min_value=0.0, format="%.2f")
            
            # Automatically calculate total
            total = q1 + q2 + q3 + q4 + q5

            # Display calculated total
            st.write(f"Total Marks for {first_name}: **{total:.2f}**")

            # Submit button
            submit_button = st.form_submit_button(label="Submit")
            if submit_button:
                student_data.append({
                    "S. No.": i + 1,
                    "Enrollment_No": enrollment_no,
                    "FirstName": first_name,
                    "Q.1": q1,
                    "Q.2": q2,
                    "Q.3": q3,
                    "Q.4": q4,
                    "Q.5": q5,
                    "Total": total
                })
                
                # Provide feedback
                st.success(f"Data for {first_name} added successfully!")

    return pd.DataFrame(student_data)

# Function to handle student data input from an uploaded Excel file
def get_student_data_from_excel(excel_file):
    df = pd.read_excel(excel_file, header=1)  # Read starting from the correct header row

    # Remove any unnamed columns
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

    # Show columns in the uploaded file for debugging
    st.write("Uploaded Excel columns:", df.columns.tolist())

    # Validate necessary columns
    required_columns = ["Enrollment_No", "FirstName", "Q.1", "Q.2", "Q.3", "Q.4", "Q.5"]
    
    # Check if required columns are present in the DataFrame
    if not all(col in df.columns for col in required_columns):
        st.error(f"Excel file must contain the following columns: {', '.join(required_columns)}")
        return pd.DataFrame()  # Return an empty DataFrame

    # Calculate total marks for each student if not present in the uploaded data
    if 'Total' not in df.columns:
        df['Total'] = df[['Q.1', 'Q.2', 'Q.3', 'Q.4', 'Q.5']].sum(axis=1)
    
    # Ensure 'S. No.' column is created
    df['S. No.'] = df.index + 1  # Adding serial number

    return df

# Function to calculate total number of students and marks
def calculate_totals(student_data):
    # Consider students as appeared if they have marks > 0 in any of the Q.1 to Q.5 columns
    appeared_students = student_data[(student_data['Q.1'] >= 0) | (student_data['Q.2'] >= 0) | 
                                     (student_data['Q.3'] >= 0) | (student_data['Q.4'] >= 0) | 
                                     (student_data['Q.5'] >= 0)]

    # Count all students who appeared in the exam
    total_students = appeared_students.shape[0]

    # Count students who attempted each question (marks > 0)
    attempted_q1 = appeared_students[appeared_students['Q.1'] > 0].shape[0]
    attempted_q2 = appeared_students[appeared_students['Q.2'] > 0].shape[0]
    attempted_q3 = appeared_students[appeared_students['Q.3'] > 0].shape[0]
    attempted_q4 = appeared_students[appeared_students['Q.4'] > 0].shape[0]
    attempted_q5 = appeared_students[appeared_students['Q.5'] > 0].shape[0]

    # Calculate total marks for each question (including zeros for non-attempts)
    q1_total = appeared_students['Q.1'].fillna(0).sum()
    q2_total = appeared_students['Q.2'].fillna(0).sum()
    q3_total = appeared_students['Q.3'].fillna(0).sum()
    q4_total = appeared_students['Q.4'].fillna(0).sum()
    q5_total = appeared_students['Q.5'].fillna(0).sum()

    # Calculate the total of "Total" marks, filling NaN with 0 if needed
    total_marks = appeared_students['Total'].fillna(0).sum()

    return {
        "Total Students Appeared": total_students,
        "Total Students Who Attempted Q.1": attempted_q1,
        "Total Students Who Attempted Q.2": attempted_q2,
        "Total Students Who Attempted Q.3": attempted_q3,
        "Total Students Who Attempted Q.4": attempted_q4,
        "Total Students Who Attempted Q.5": attempted_q5,
        "Overall Total Marks": total_marks
    }

# Function to generate the CO vs. Questions table
def generate_co_question_table(co_marks, question_data):
    # Create a dictionary to store CO vs Question mapping with marks
    co_question_data = defaultdict(lambda: [0] * len(question_data))  # Create a list of zeros for each question

    for qdata in question_data:
        question_index = qdata['question_number'] - 1  # Zero-indexed for list storage
        cos = qdata['cos']
        marks = qdata['marks']

        # Distribute total marks to each CO, without splitting
        for co in cos:
            co_question_data[co][question_index] = marks  # Assign full marks to each CO

    # Convert the dictionary to a DataFrame
    co_df = pd.DataFrame.from_dict(co_question_data, orient='index', columns=[f"Q.{i+1}" for i in range(len(question_data))])

    # Add COs as Enrollment_No column
    co_df.reset_index(inplace=True)
    co_df.rename(columns={'index': 'CO'}, inplace=True)

    # Define the order of COs
    co_order = ['CO1', 'CO2', 'CO3', 'CO4']

    # Reindex the DataFrame to include all COs in the defined order, filling missing rows with zeros
    co_df = co_df.set_index('CO').reindex(co_order, fill_value=0).reset_index()

    return co_df


#function to generate MARKS * APPEARED table
def generate_student_co_table(co_marks, question_data, students_attempted):
    # Create a dictionary to store CO vs Question mapping with marks
    student_co_data = defaultdict(lambda: [0] * len(question_data))  # Create a list of zeros for each question

    for qdata in question_data:
        question_index = qdata['question_number'] - 1  # Zero-indexed for list storage
        cos = qdata['cos']
        marks = qdata['marks']

        # Get the number of students who attempted this question
        no_of_students_attempted = students_attempted[question_index] if question_index < len(students_attempted) else 0

        # Calculate total marks for the question based on the number of students who attempted it
        total_marks_for_question = marks * no_of_students_attempted  # Multiply marks with the number of students

        # Assign total marks to each CO without splitting
        for co in cos:
            student_co_data[co][question_index] = total_marks_for_question  # Assign full marks for each CO

    # Convert the dictionary to a DataFrame
    co_df = pd.DataFrame.from_dict(student_co_data, orient='index', columns=[f"Q.{i+1}" for i in range(len(question_data))])

    # Add COs as Enrollment_No column
    co_df.reset_index(inplace=True)
    co_df.rename(columns={'index': 'CO'}, inplace=True)

    # Define the order of COs
    co_order = ['CO1', 'CO2', 'CO3', 'CO4']

    # Reindex the DataFrame to include all COs in the defined order, filling missing rows with zeros
    co_df = co_df.set_index('CO').reindex(co_order, fill_value=0).reset_index()

    return co_df


# Function to generate the CO metrics table
def generate_co_metrics_table(co_marks, student_co_df, total_students):
    co_order = ['CO1', 'CO2', 'CO3', 'CO4']  # Define the order of COs
    co_metrics_data = []

    for co in co_order:  # Loop through the defined order of COs
        if co in co_marks:
            # Sum of (marks * number of students appeared in that CO)
            sum_marks_students = student_co_df.loc[student_co_df['CO'] == co].select_dtypes(include=['float64', 'int64']).sum(axis=1).values[0]
            
            # Sum of total marks asked in that CO
            sum_marks_asked = co_marks[co]
            
            # Metric 1: (Sum of marks * no. of students appeared) / (Sum of marks asked in that CO)
            co_average = sum_marks_students / sum_marks_asked if sum_marks_asked != 0 else 0
            
            # Metric 2: (Metric 1) / Total number of students appeared
            co_normalized = co_average / total_students if total_students != 0 else 0
        else:
            # If CO is not in co_marks, fill with zeros
            sum_marks_students = 0
            co_average = 0
            co_normalized = 0

        # Add metrics for the CO to the data list
        co_metrics_data.append({
            'CO': co,
            'Sum of (Marks * No. of Students)': sum_marks_students,
            'Metric 1': co_average,
            'Metric 2': co_normalized
        })

    # Convert to DataFrame for display
    co_metrics_df = pd.DataFrame(co_metrics_data)

    return co_metrics_df


# Function to create an editable table for CO-PO mapping
def get_co_po_mapping():
    # Define COs and POs
    co_list = ['CO1', 'CO2', 'CO3', 'CO4']
    po_list = [f"PO{i}" for i in range(1, 13)]  # PO1 to PO12

    # Create an empty DataFrame to store the CO-PO mapping
    co_po_df = pd.DataFrame(index=co_list, columns=po_list)

    st.write("Please enter CO-PO mapping values for each cell (leave empty for no value).")
    
    # Create form for user inputs
    with st.form(key="co_po_form"):
        for co in co_list:
            st.write(f"## {co}")
            for po in po_list:
                # Use a text input for each CO-PO cell, with CO and PO as part of the unique key
                value = st.text_input(f"{co} - {po}", key=f"{co}_{po}")
                
                # Try to convert the input to a float, if not possible, assign NaN
                try:
                    co_po_df.loc[co, po] = float(value) if value else float("nan")
                except ValueError:
                    st.error(f"Please enter a valid number for {co} - {po}")

        # Submit button to collect inputs
        submit_button = st.form_submit_button(label="Submit CO-PO Mapping")

    # Return the DataFrame after the form is submitted
    if submit_button:
        # Calculate the average for each PO (only if all COs have a value)
        po_averages = co_po_df.mean(axis=0, skipna=True)
        po_averages = po_averages.where(co_po_df.notna().all(axis=0))

        # Calculate the average for each CO (only for present values)
        co_averages = co_po_df.mean(axis=1, skipna=True)

        # Add averages to the DataFrame
        co_po_df.loc['Average'] = po_averages
        co_po_df['Average'] = co_averages

        # Calculate the average of CO averages and set it in the last row of the Average column
        average_of_co_averages = co_averages.mean() if not co_averages.isna().all() else float("nan")
        co_po_df.at['Average', 'Average'] = average_of_co_averages

        # Show the entered CO-PO mapping
        st.write("Here is the CO-PO Mapping you entered with averages:")
        st.dataframe(co_po_df)

        # Provide option to download the CO-PO mapping as Excel (you can implement this function)
        # download_excel(co_po_df, "CO_PO_Mapping.xlsx")
        
        return co_po_df

    return None


def generate_co_po_metrics_table(co_po_mapping, co_metrics_df):
    if co_po_mapping is None:
        raise ValueError("co_po_mapping cannot be None")
    # Initialize a new DataFrame to hold the results
    co_po_metrics_df = pd.DataFrame(co_po_mapping)

    # Create an empty DataFrame to store CO-PO metrics results
    co_po_results = pd.DataFrame(index=co_po_mapping.index, columns=co_po_mapping.columns)

    # Iterate over each CO to calculate the multiplied values
    for co in co_po_mapping.index:
        for po in co_po_mapping.columns:
            if not pd.isna(co_po_mapping.at[co, po]) and co in co_metrics_df['CO'].values:
                # Get Metric 2 for the corresponding CO
                metric_2_value = co_metrics_df.loc[co_metrics_df['CO'] == co, 'Metric 2'].values[0]
                co_po_results.at[co, po] = co_po_mapping.at[co, po] * metric_2_value

    # Calculate the average for each PO (only if all COs have a value)
    po_averages = co_po_results.mean(axis=0, skipna=True)

    # Calculate the average for each CO (only for present values)
    co_averages = co_po_results.mean(axis=1, skipna=True)

    # Add averages to the results DataFrame
    co_po_results.loc['Average'] = po_averages
    co_po_results['Average'] = co_averages

    # Calculate the average of CO averages and set it in the last row of the Average column
    average_of_co_averages = co_averages.mean() if not co_averages.isna().all() else float("nan")
    co_po_results.at['Average', 'Average'] = average_of_co_averages

    return co_po_results

def average_of_co_averages(df):
    if 'Average' in df.columns:
        return df['Average'].mean() if not df['Average'].isna().all() else float("nan")
    return float("nan")


def calculate_attainment(co_po_metrics_df, co_po_mapping_df):
    avg_metrics = average_of_co_averages(co_po_metrics_df)
    avg_mapping = average_of_co_averages(co_po_mapping_df)

    if avg_mapping > 0:  # Prevent division by zero
        attainment = (avg_metrics / avg_mapping) * 100
    else:
        attainment = 0

    return attainment

def main():
    st.title("CO Marks Extractor")

    # Upload PDF
    pdf_file = st.file_uploader("Upload PDF", type="pdf")
    if pdf_file is not None:
        text = extract_text_from_pdf(pdf_file)
        question_data = extract_question_data(text)
        co_marks = aggregate_marks_by_co(question_data)

        # Display question data
        if co_marks:
            st.subheader("Extracted Questions and CO Marks")
            for qdata in question_data:
                st.write(f"Q{qdata['question_number']}: COs: {', '.join(qdata['cos'])}, Marks: {qdata['marks']}")

            st.subheader("Total Marks per CO:")
            for co, marks in co_marks.items():
                st.write(f"{co}: {marks} marks")

            # Plotting pie chart
            labels = co_marks.keys()
            sizes = co_marks.values()
            fig, ax = plt.subplots()
            ax.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=140)
            ax.axis('equal')  # Equal aspect ratio ensures that pie chart is circular
            st.pyplot(fig)

        else:
            st.info("No COs or marks detected.")

    # Option to enter student data manually or upload Excel
    st.subheader("Student Data Input")
    input_method = st.radio("Choose Input Method", ("Manual Input", "Upload Excel"))

    # Initialize student_data
    student_data = pd.DataFrame()  # Ensure student_data is initialized as a DataFrame

    if input_method == "Manual Input":
        student_data = get_student_data_manually()
    else:
        excel_file = st.file_uploader("Upload Excel", type=["xlsx", "xls"])
        if excel_file is not None:
            student_data = get_student_data_from_excel(excel_file)
            if not student_data.empty:
                st.subheader("Student Data Extracted from Excel")
                st.write(student_data)

    if not student_data.empty:  # Use .empty to check if the DataFrame is empty
        st.subheader("Student Data Entered")
        st.write(student_data)

        # Calculate totals
        totals = calculate_totals(student_data)

        # Display totals
        st.subheader("Summary of Student Marks")
        for key, value in totals.items():
            st.write(f"{key}: {value}")

        # Modify the first name to show total number of students
        total_students_text = f"(Total No. of Students = {totals['Total Students Appeared']})"
        student_data.at[student_data.index[-1], 'FirstName'] = total_students_text

        # Add total marks and total students appeared per question in respective columns
        student_data.at[student_data.index[-1], 'Q.1'] = totals['Total Students Who Attempted Q.1']
        student_data.at[student_data.index[-1], 'Q.2'] = totals['Total Students Who Attempted Q.2']
        student_data.at[student_data.index[-1], 'Q.3'] = totals['Total Students Who Attempted Q.3']
        student_data.at[student_data.index[-1], 'Q.4'] = totals['Total Students Who Attempted Q.4']
        student_data.at[student_data.index[-1], 'Q.5'] = totals['Total Students Who Attempted Q.5']
        student_data.at[student_data.index[-1], 'Total'] = totals['Overall Total Marks']

        # Display modified student data
        st.subheader("Updated Student Data with Totals")
        st.write(student_data)

        # Generate and display CO vs Questions table
        co_question_df = generate_co_question_table(co_marks, question_data)
        st.subheader("CO vs Questions Table")
        st.write(co_question_df)

        # Generate and display Student CO Table
        student_co_df = generate_student_co_table(co_marks, question_data, 
                                                   [totals['Total Students Who Attempted Q.1'],
                                                    totals['Total Students Who Attempted Q.2'],
                                                    totals['Total Students Who Attempted Q.3'],
                                                    totals['Total Students Who Attempted Q.4'],
                                                    totals['Total Students Who Attempted Q.5']])
        st.subheader("Student CO Marks Table")
        st.write(student_co_df)

        # Generate and display the new CO metrics table
        co_metrics_df = generate_co_metrics_table(co_marks, student_co_df, totals['Total Students Appeared'])
        st.subheader("CO Metrics Table")
        st.write(co_metrics_df)




        st.title("CO-PO Mapping Input")

        # Call the function to get CO-PO mapping from the user
        co_po_mapping = get_co_po_mapping()

        # After the user inputs the data, show the DataFrame
        if co_po_mapping is not None:
            st.write("CO-PO Mapping has been entered successfully.")

         # Generate and display the CO Metrics Table
        co_metrics_df = generate_co_metrics_table(co_marks, student_co_df, totals['Total Students Appeared'])
        st.subheader("CO Metrics Table")
        st.write(co_metrics_df)

        
         # Generate and display CO-PO Metrics Table
        if co_po_mapping is not None:
            co_po_metrics_df = generate_co_po_metrics_table(co_po_mapping, co_metrics_df)
            st.subheader("CO-PO Metrics Table")
            st.write(co_po_metrics_df)

        # Calculate attainment
        attainment = calculate_attainment(co_po_metrics_df, co_po_mapping)
        st.subheader("Attainment")
        st.write(f"The calculated attainment is: {attainment:.2f}%")    

        # # Option to download the updated student data as Excel
        # excel_file_name = "Updated_Student_Data.xlsx"
        # if st.button("Download Updated Student Data"):
        #     with pd.ExcelWriter(excel_file_name) as writer:
        #         student_data.to_excel(writer, index=False)
        #     st.success(f"{excel_file_name} has been created successfully!")

if __name__ == "__main__":
    main()