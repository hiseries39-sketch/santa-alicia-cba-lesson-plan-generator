import streamlit as st, fitz, io, datetime, requests, re
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from urllib.parse import quote

st.set_page_config(page_title='SANTA ALICIA CBA LESSON PLAN GENERATOR', layout='wide', page_icon='🏫')
st.title('🏫 SANTA ALICIA CBA LESSON PLAN GENERATOR')
st.markdown('**SANTA ALICIA PRESTIGE BILINGUAL COLLEGE - Official CBA Teacher Toolkit | Form 1 to Upper 6th | CS & ICT**')

def search_auto(topic):
    try:
        url = f"https://en.wikipedia.org/api/rest_v1/page/summary/{quote(topic)}"
        r = requests.get(url, timeout=6)
        if r.status_code == 200:
            data = r.json()
            return data.get('extract', ''), data.get('content_urls',{}).get('desktop',{}).get('page','')
    except: pass
    return f"{topic} is a key concept in Cameroon GCE Computer Science/ICT syllabus. It involves practical application in real life.", ""

# ===== CURRICULUM DEFINITIONS =====
FORM_1_4_CS = {
    'MODULE 1: INTRODUCTION TO COMPUTERS': ['Definition of Computer, Uses', 'Characteristics, Advantages/Disadvantages'],
    'MODULE 2: COMPUTER HARDWARE': ['Input/Output Devices', 'System Unit, Storage'],
    'MODULE 3: COMPUTER SOFTWARE': ['System Software vs Application', 'Examples of Software'],
    'MODULE 4: BASIC ICT SKILLS': ['Keyboard, Mouse Skills', 'File Creation, Saving'],
}

FORM5_CS = {
    'MODULE 1: HISTORY AND EVOLUTION OF COMPUTERS': ['Generations of Computers', 'Contributions of Babbage, Hollerith', 'Types of Computers'],
    'MODULE 2: COMPUTER HARDWARE': ['Input Devices, Output Devices', 'CPU Components ALU CU', 'Primary & Secondary Storage'],
    'MODULE 3: NUMBER SYSTEM': ['Binary, Decimal, Octal, Hex Conversions', 'Binary Addition, Subtraction, One Two Complement'],
    'MODULE 4: LOGIC GATE AND BOOLEAN ALGEBRA': ['Basic Gates AND OR NOT', 'Universal NAND NOR, Truth Tables', 'Boolean Laws, Simplification, K-Map'],
    'MODULE 5: FILE FORMAT AND FILE MANAGEMENT': ['File Formats .txt .docx .mp3 .jpg', 'File System FAT NTFS, File Management Operations'],
    'MODULE 6: COMPUTER ERGONOMICS AND COMPUTER ETHICS': ['Ergonomics, Health Issues', 'Computer Ethics, Cyber Crimes, Copyright, Intellectual Property'],
    'MODULE 7: COMPUTER SOFTWARE': ['System Software Types', 'Application Software, Utility, Open vs Proprietary'],
    'MODULE 8: OPERATING SYSTEM': ['Functions of OS, Types of OS', 'File Management, Memory Management', 'Booting Process, OS Interface'],
    'MODULE 9: DATA STRUCTURE': ['Arrays 1D 2D', 'Stacks LIFO, Queues FIFO', 'Linked Lists Types'],
    'MODULE 10: DATABASE': ['DBMS Definition, Advantages', 'Keys Primary Foreign, ER Diagram', 'SQL Basics'],
    'MODULE 11: COMPUTER NETWORKING AND DATA COMMUNICATION': ['Types of Network LAN WAN MAN', 'Network Topologies, Devices', 'OSI Model 7 Layers, IP Address, Internet'],
    'MODULE 12: BASIC CONCEPTS OF PROJECT MANAGEMENT': ['Project Definition, WBS, Gantt Chart', 'Project Life Cycle'],
    'MODULE 13: ALGORITHMS': ['Definition, Flowchart Symbols', 'Pseudocode, Sorting, Searching'],
    'MODULE 14: PROBLEM SOLVING AND CODING': ['Problem Solving Steps, Debugging', 'Testing Techniques'],
    'MODULE 15: PROGRAMMING': ['Variables, Data Types, Control Structures IF Loop', 'Functions, Python/C Practical, GCE Questions']
}

CS_SIXTH = {
    'MODULE 1: COMPUTER SYSTEM AND ARCHITECTURE': ['CPU Architecture Von Neumann', 'Memory Hierarchy Cache RAM', 'Fetch Decode Execute Cycle, Buses'],
    'MODULE 2: COMPUTER SOFTWARE': ['System vs Application Software', 'Software Development Life Cycle Models'],
    'MODULE 3: OPERATING SYSTEM': ['Process Management Scheduling', 'Memory Management Paging', 'Deadlock, File Management'],
    'MODULE 4: DATA STRUCTURE': ['Stacks Queues Implementation', 'Linked Lists Types', 'Trees Binary, Graphs BFS DFS'],
    'MODULE 5: DATABASE': ['Relational Model, ERD Extended', 'Normalization 1NF 2NF 3NF', 'SQL DDL DML Advanced'],
    'MODULE 6: COMPUTER NETWORKING AND DATA COMMUNICATION': ['OSI vs TCP/IP Model', 'Subnetting CIDR, Routing', 'Network Security'],
    'MODULE 7: BASIC CONCEPTS OF PROJECT MANAGEMENT': ['WBS, Gantt, PERT CPM', 'Risk Management'],
    'MODULE 8: ALGORITHMS': ['Sorting Quick Sort Merge Sort Complexity', 'Searching Binary, Algorithm Complexity Big O'],
    'MODULE 9: PROBLEM SOLVING AND CODING': ['Debugging Techniques', 'Testing Unit Integration'],
    'MODULE 10: PROGRAMMING': ['OOP Class Object Inheritance Polymorphism', 'File Handling, Exception Handling, Project']
}

ICT_SIXTH = {
    'MODULE 1: CAREER INFORMATION TECHNOLOGY AND COMPUTER APPLICATION': ['ICT Careers, Job Roles', 'Emerging Tech AI IoT Cloud, Computer Applications Word Excel'],
    'MODULE 2: COMPUTER ERGONOMICS AND COMPUTER ETHICS': ['Ergonomics Principles, Health Safety', 'Ethics, Cyber Crime, Copyright, Privacy'],
    'MODULE 3: COMPUTER SOFTWARE': ['System Software, Application Software', 'Software Acquisition Licensing'],
    'MODULE 4: OPERATING SYSTEM': ['Functions of OS, Types', 'Windows Linux Operations'],
    'MODULE 5: COMPUTER NETWORK AND DATA COMMUNICATION': ['Network Types LAN WAN, Topologies', 'OSI Model, TCP/IP, IP Address, Internet Services'],
    'MODULE 6: INFORMATION SYSTEM AND DESIGN': ['SDLC, Information System Types MIS DSS', 'System Design ERD DFD, Database Design'],
    'MODULE 7: BASIC CONCEPTS OF PROJECT MANAGEMENT': ['Project Definition, Planning WBS Gantt', 'Project Execution'],
    'MODULE 8: ALGORITHMS': ['Algorithm Definition, Flowchart', 'Sorting Searching'],
    'MODULE 9: PROGRAMMING': ['Introduction to Programming, Variables Control Structures', 'OOP Basics, Web Programming Basics']
}

def get_curr(level, subj):
    if level in ['Form 1','Form 2','Form 3','Form 4']: return FORM_1_4_CS
    if level == 'Form 5': return FORM5_CS
    if level in ['Lower 6th','Upper 6th']:
        return CS_SIXTH if subj == 'Computer Science' else ICT_SIXTH
    return FORM5_CS

# ===== SIDEBAR =====
st.sidebar.title('SANTA ALICIA CBA')
subject = st.sidebar.selectbox('Subject', ['Computer Science','ICT'])
level = st.sidebar.selectbox('Class / Level', ['Form 1','Form 2','Form 3','Form 4','Form 5','Lower 6th','Upper 6th'], index=4)
st.sidebar.markdown('---')
st.sidebar.info('**SANTA ALICIA CBA LESSON PLAN GENERATOR**\nAuto-writes by topic\n100 Minutes Plan')

tab1, tab2 = st.tabs(['📚 CBA LESSON PLAN GENERATOR (Auto-Writes)', '📄 GCE OL/AL PAPER 1,2,3 PDF TO WORD'])

with tab1:
    st.subheader(f'📚 {level} | {subject} | SANTA ALICIA CBA LESSON PLAN GENERATOR')
    curr = get_curr(level, subject)
    mod = st.selectbox(f'📚 MODULE / STRAND ({len(curr)} Modules)', list(curr.keys()))
    lesson = st.selectbox(f'📖 SUB-STRAND / TOPIC ({len(curr[mod])} Lessons)', curr[mod])
    
    colA, colB = st.columns(2)
    with colA:
        topic_input = st.text_input('✏️ ENTER TOPIC (AI will auto-search & write)', value=lesson, placeholder='e.g. Binary Addition or Types of Network')
        school = st.text_input('School / Institution Name', 'SANTA ALICIA PRESTIGE BILINGUAL COLLEGE')
        teacher = st.text_input('Teacher Name', '')
        date = st.date_input('Date', datetime.date.today())
    with colB:
        time = st.text_input('Time & Duration', '8:00 - 9:40 AM (100 Minutes)')
        boys = st.number_input('Boys', 0, 100, 20)
        girls = st.number_input('Girls', 0, 100, 25)
        total = boys+girls
        ref = st.text_input('References', 'Cameroon GCE Syllabus, Computer Science Textbook Form 5, Teacher Guide, Internet')

    real_input = st.text_area('Real-Life Situation (Context) - Optional AI will generate if empty', '', placeholder='e.g. An IT technician at ENEO generating bills via mass-command...')
    aids = st.text_input('Teaching and Learning Aids / Resources', 'Chalkboard, Markers, Computer, Projector, Flashcards, Handouts, Internet')

    if st.button('✨ GENERATE CBA LESSON PLAN - 100 MINS - AUTO WRITES', type='primary', use_container_width=True):
        final_topic = topic_input.strip() if topic_input.strip() else lesson
        with st.spinner(f'AI Searching {final_topic} and writing 100 mins CBA plan...'):
            auto_def, auto_link = search_auto(final_topic)
            
            # Auto generate context if empty
            if not real_input.strip():
                real_situation = f"In the community of Bamenda, a local business / school / hospital needs to use {final_topic} to solve a daily problem. For example, as an IT technician / database manager / network administrator, learners must apply {final_topic} to improve efficiency, similar to ENEO billing system or school result management at {school}."
            else:
                real_situation = real_input

            # ===== OFFICIAL CBA FORMAT =====
            plan_text = f"""
SANTA ALICIA PRESTIGE BILINGUAL COLLEGE
SANTA ALICIA CBA LESSON PLAN GENERATOR - OFFICIAL CBA FORMAT - 100 MINUTES

1. ADMINISTRATIVE & IDENTIFICATION DETAILS (PREAMBLE)

School / Institution Name: {school}
Class / Level: {level}
Subject / Learning Area: {subject}
Module / Strand: {mod}
Sub-strand / Topic: {final_topic}
Date: {date} | Time: {time} | Duration: 100 Minutes
Class Size / Attendance: Total {total} (Boys: {boys}, Girls: {girls}) | Present: {total}
References: {ref}
Teacher: {teacher}
Source Auto-Searched: {auto_link}
Auto Definition: {auto_def}

2. CONTEXT & COMPETENCY FOUNDATION

A. Real-Life Situation (Context):
{real_situation}

B. Core Competencies / Life Skills:
- Critical Thinking & Problem Solving
- Effective Communication & Collaboration
- Digital Literacy & ICT Skills
- Creativity & Innovation
- Citizenship & Ethics

C. Expected Learning Outcomes / Competencies:
At the end of 100 minutes, learners can / is able to:
1. Define and explain {final_topic} correctly with examples.
2. Recall and describe types / components / functions of {final_topic}.
3. Apply {final_topic} to solve the real-life situation of {real_situation[:80]}...
4. Demonstrate practical skills in {final_topic} through group activity.
5. Exhibit ethical use of {final_topic} in society.

D. Teaching and Learning Aids / Resources:
{aids}, Whiteboard, Lesson Note on {final_topic}, GCE Past Questions.

3. LESSON DEVELOPMENT STAGES (PROCEDURE) - 100 MINUTES TOTAL


PHASE 1: INTRODUCTION / STARTER (15 Minutes)
Time: 0:00 - 0:15
- Review Previous Knowledge (5 min): Teacher asks: "What did we learn last lesson on {mod}?" Learners recall. Q&A on {lesson}.
- Hook / Entry Behaviour (5 min): Present Real-Life Problem: "{real_situation}" - Ask learners how they would solve it without computers.
- Announcement (5 min): Teacher states: "Today we learn {final_topic}. By end you can solve above problem." Write topic on board. Share Learning Outcomes. Link to GCE relevance.

PHASE 2: LESSON DEVELOPMENT / ACTIVITIES (70 Minutes - Main Body)

STEP 1: CONCEPT DEVELOPMENT - Definition & Importance (25 Minutes) 0:15 - 0:40
Teacher Role (Facilitator): Explains definition using auto-searched content: "{auto_def}" Uses chalkboard, projector to illustrate. Asks probing questions.
Learner Role (Active): Listen, take notes, ask questions, give examples of {final_topic} in daily life (phone, school, market).
Method: Interactive Lecture, Questioning, Brainstorming, Demonstration
Pacing: 10 min explanation, 10 min examples, 5 min Q&A

STEP 2: TYPES / COMPONENTS / PRACTICAL ANALYSIS (25 Minutes) 0:40 - 1:05
Teacher Role: Divides class into 4 groups of {total//4} learners. Gives each group task: Group 1 - Identify types of {final_topic}, Group 2 - Functions, Group 3 - Advantages/Disadvantages, Group 4 - Real application to context.
Learner Role: Group work, discussion, research from textbook/handout, present findings, peer teaching.
Method: Collaborative Learning, Guided Discovery, Inquiry-Based Learning, Think-Pair-Share
Pacing: 5 min grouping, 10 min discussion, 10 min presentation
Teaching Aid: Flashcards, Chart showing {final_topic}

STEP 3: APPLICATION / PROBLEM SOLVING / HANDS-ON (20 Minutes) 1:05 - 1:25
Teacher Role: Gives practical task: "Using {final_topic}, solve: {real_situation} Write steps/algorithm/flowchart/practical on computer."
Learner Role: Problem solver, applies knowledge, does practical, demonstrates.
Method: Project-Based, Hands-on Practical, Problem Solving, Demonstration
Pacing: 5 min instruction, 10 min practical, 5 min demo by 2 learners
Competency: Digital Literacy, Problem Solving

PHASE 3: EVALUATION / ASSESSMENT (10 Minutes) 1:25 - 1:35
Practical Tasks to Test Mastery:
1. Define {final_topic} (Knowledge - 4 marks)
2. List 3 types / components of {final_topic} (Knowledge - 6 marks)
3. Explain how {final_topic} solves: "{real_situation[:100]}" (Application - 6 marks)
4. Practical: Draw / Write / Demonstrate {final_topic} (Experimental - 4 marks)
Total: 20 Marks
Method: Formative Assessment, Peer Marking (exchange books), Self-Assessment, Teacher corrects using GCE marking scheme.
GCE Link: This evaluation matches GCE {subject} Paper 2 structure.

PHASE 4: PLENARY / CLOSURE & EXTENDED ACTIVITY (5 Minutes) 1:35 - 1:40
- Summary / Recap (3 min): Learner-led recap: "What did we learn today on {final_topic}?" Teacher reinforces key points, link back to real-life situation - now learners can solve it.
- Extended Activity / Homework (2 min): 
  Research: Find 2 more real-life uses of {final_topic} in Bamenda.
  GCE Past Question: Answer GCE {level} {subject} question on {final_topic}.
  Practical: Practice {final_topic} on computer/phone at home.
- Values: Remind ethics, ergonomics, teamwork.

4. TEACHER REFLECTION / REMARKS

Lesson was delivered successfully within 100 minutes. Learners participated actively. {total} learners present. Objectives achieved. Next lesson: Continue with next sub-topic in {mod}.

Generated by: SANTA ALICIA CBA LESSON PLAN GENERATOR | {school} | {date}
Auto-Searched Definition: {auto_def[:300]}
"""
            st.success(f'✅ 100 Mins CBA Plan Auto-Written for {final_topic} - {level}')
            st.text_area('📄 OFFICIAL CBA PLAN - 100 MINS - PREVIEW', plan_text, height=600)
            
            # DOCX EXPORT
            doc = Document()
            style = doc.styles['Normal']
            style.font.name = 'Times New Roman'
            style.font.size = Pt(12)
            for para in plan_text.split('\n'):
                p = doc.add_paragraph(para)
                if '====' in para or 'ADMINISTRATIVE' in para or 'CONTEXT' in para or 'LESSON DEVELOPMENT' in para or 'PHASE' in para:
                    for run in p.runs:
                        run.bold = True
            out = io.BytesIO()
            doc.save(out)
            st.download_button('📥 DOWNLOAD OFFICIAL CBA PLAN - MS WORD (.docx) - 100 MINS', out.getvalue(), file_name=f'SANTA_ALICIA_CBA_LESSON_PLAN_GENERATOR_{level}_{final_topic[:25].replace(" ","_")}_100MINS.docx', type='primary')

with tab2:
    st.header('📄 GCE OL/AL PAPER 1,2,3 - PDF TO FORMATTED WORD - SANTA ALICIA CBA LESSON PLAN GENERATOR')
    st.info('Search Cameroon GCE Question Papers PDF and Convert to MS Word with respect to Paper 1, Paper 2, Paper 3 Formats')
    c1,c2,c3,c4 = st.columns(4)
    with c1: gce_level = st.selectbox('GCE LEVEL', ['GCE O-Level (Form 5)','GCE A-Level (6th Form)'])
    with c2: paper = st.selectbox('PAPER FORMAT', ['Paper 1 - MCQ (40 Questions)','Paper 2 - Theory/Essay','Paper 3 - Practical/Alternative'])
    with c3: year = st.selectbox('YEAR', list(range(2025,2010,-1)))
    with c4: subj_p = st.selectbox('SUBJECT', ['Computer Science','ICT'])
    
    if st.button('🔍 SEARCH GCE PAPERS PDF BY PAPER FORMAT'):
        q = f"Cameroon GCE {gce_level} {subj_p} {paper} {year} PDF camgcebible past questions"
        st.markdown(f"""
        **Search Results for {gce_level} {subj_p} {paper} {year}:**
        - 👉 [Google Search: {gce_level} {paper} {year}](https://www.google.com/search?q={quote(q)})
        - 👉 [Search camgcebible.com](https://www.google.com/search?q=site:camgcebible.com+{subj_p}+{paper}+{year})
        - 👉 [Search GCE Past Papers PDF](https://www.google.com/search?q=Cameroon+GCE+{subj_p}+{year}+past+questions+PDF+download)
        """)
    
    st.divider()
    pdf_file = st.file_uploader(f'Upload GCE {gce_level} {paper} PDF', type=['pdf'])
    if pdf_file:
        if st.button('🚀 CONVERT TO FORMATTED MS WORD - PAPER 1,2,3 FORMAT', type='primary', use_container_width=True):
            with st.spinner('Converting to Formatted Word - Preserving Paper 1,2,3 Format...'):
                pdf_doc = fitz.open(stream=pdf_file.read(), filetype='pdf')
                raw = "\n".join([p.get_text() for p in pdf_doc])
                doc = Document()
                style = doc.styles['Normal']; style.font.name='Times New Roman'; style.font.size=Pt(12)
                header = doc.add_paragraph(); header.alignment=WD_ALIGN_PARAGRAPH.CENTER
                run = header.add_run(f'CAMEROON GENERAL CERTIFICATE OF EDUCATION BOARD\n{gce_level.upper()} EXAMINATION\n{subj_p.upper()} - {paper.upper()} - {year}\nSANTA ALICIA CBA LESSON PLAN GENERATOR - STUDY COPY\n\n'); run.bold=True; run.font.size=Pt(13)
                
                if 'Paper 1' in paper:
                    doc.add_paragraph('INSTRUCTIONS: PAPER 1 - MULTIPLE CHOICE OBJECTIVE\nDuration: 1 Hour 30 Minutes\nAnswer ALL 40 Questions. Each question carries 1 mark. Choose A, B, C, D.\n').bold=True
                    qs = re.split(r'\n(?=\d{1,2}[\.\)])', raw)
                    for q in qs:
                        if q.strip(): doc.add_paragraph(q.strip())
                elif 'Paper 2' in paper:
                    doc.add_paragraph('INSTRUCTIONS: PAPER 2 - THEORY / ESSAY\nDuration: 2 Hours 30 Min (O-Level), 3 Hours (A-Level)\nAnswer ANY FIVE Questions. All questions carry equal marks.\n').bold=True
                    for line in raw.split('\n'):
                        p = doc.add_paragraph(line)
                        if re.match(r'^\s*\d+[\.\)]', line): 
                            for r in p.runs: r.bold=True
                else:
                    doc.add_paragraph('INSTRUCTIONS: PAPER 3 - PRACTICAL / ALTERNATIVE TO PRACTICAL\nDuration: 2 Hours\nAnswer ALL Questions. Follow instructions carefully.\n').bold=True
                    for line in raw.split('\n'):
                        doc.add_paragraph(line)
                
                foot = doc.add_paragraph(); foot.alignment=WD_ALIGN_PARAGRAPH.CENTER
                foot.add_run(f'\n--- END OF {paper} {year} ---\nFormatted by SANTA ALICIA CBA LESSON PLAN GENERATOR').italic=True
                out = io.BytesIO(); doc.save(out)
                st.success(f'✅ Converted {len(pdf_doc)} pages - Formatted as {paper} - {gce_level}')
                st.text_area('Preview', raw[:1500], height=200)
                st.download_button('📥 DOWNLOAD FORMATTED WORD - PAPER 1,2,3 READY', out.getvalue(), file_name=f'SANTA_ALICIA_CBA_{gce_level.replace(" ","_")}_{subj_p}_{paper.replace(" ","_")}_{year}_FORMATTED.docx', type='primary')