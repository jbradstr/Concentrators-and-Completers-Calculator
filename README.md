# Concentrators-and-Completers-Calculator-
This is a macro I created that a data analyst at a school district uses to go from raw data pulled from PowerSchool to data that can be quickly uploaded back to PowerSchool.  The macro takes the courses that a student took and increments them specifically to what the requirements are for a concentrator or a completer within each program of study.  Once the student has been incremented there are checks after each incrementation that will display information based on when they became a concetrator or a completer.   

```vbscript

Public Sub increments()

    'check to see if kids ever take the same course twice even if they pass it the first time

    Dim main As Worksheet: Set main = ThisWorkbook.Worksheets("main")
    Dim pc As Worksheet: Set pc = ThisWorkbook.Worksheets("Programs_Courses")
    Dim pso As Worksheet: Set pso = ThisWorkbook.Worksheets("ps_output")
    lastrow = main.Range("A" & main.Rows.Count).End(xlUp).Row
    Dim DestRow As Integer
   
    'programs variable declarations
    Dim acc_con As Integer, _
    acc_comp As Integer, _
    pl_an_sys_con As Integer, _
    pl_an_sys_comp As Integer, _
    auto_tech_con As Integer, _
    auto_tech_comp As Integer, _
    cul_art_con As Integer, _
    cul_art_comp As Integer, _
    dig_art_des_con As Integer, _
    dig_art_des_comp As Integer, _
    ece_con As Integer, _
    ece_comp As Integer, _
    envi_nat_res_con As Integer, _
    envi_nat_res_comp As Integer

    Dim gen_manag_con As Integer, _
    gen_manag_comp As Integer, _
    graph_comm_con As Integer, _
    graph_comm_comp As Integer, _
    health_sci_con As Integer, _
    health_sci_comp As Integer, _
    horti_con As Integer, _
    horti_comp As Integer, _
    mark_manag_con As Integer, _
    mark_manag_comp As Integer, _
    media_tech_con As Integer, _
    media_tech_comp As Integer, _
    ops_manag_con As Integer, _
    ops_manag_comp As Integer, _
    pltw_biomed_sci_con As Integer, _
    pltw_biomed_sci_comp As Integer, _
    pltw_pre_eng_con As Integer, _
    pltw_pre_eng_comp As Integer, _
    sport_med_con As Integer, _
    sport_med_comp As Integer, _
    fam_con_sci_con As Integer, _
    fam_con_sci_comp As Integer, _
    prog_soft_dev_con As Integer, _
    prog_soft_dev_comp As Integer
   
   
    'initialize program variables to 0
    acc_con = 0
    acc_comp = 0
    pl_an_sys_con = 0
    pl_an_sys_comp = 0
    auto_tech_con = 0
    auto_tech_comp = 0
    cul_art_con = 0
    cul_art_comp = 0
    dig_art_des_con = 0
    dig_art_des_comp = 0
    ece_con = 0
    ece_comp = 0
    envi_nat_res_con = 0
    envi_nat_res_comp = 0
    gen_manag_con = 0
    gen_manag_comp = 0
    graph_comm_con = 0
    graph_comm_comp = 0
    health_sci_con = 0
    health_sci_comp = 0
    horti_con = 0
    horti_comp = 0
    mark_manag_con = 0
    mark_manag_comp = 0
    media_tech_con = 0
    media_tech_comp = 0
    pltw_biomed_sci_con = 0
    pltw_biomed_sci_comp = 0
    pltw_pre_eng_con = 0
    pltw_pre_eng_comp = 0
    sports_med_con = 0
    sports_med_comp = 0
    fam_con_sci_con = 0
    fam_con_sci_comp = 0
    prog_soft_dev_con = 0
    prog_soft_dev_comp = 0
    DestRow = 2
   
   
    For i = 2 To lastrow
       
        'if student 1 is equal to student 2
        If main.Cells(i, 1) = main.Cells(i + 1, 1) Then
           
            'Accounting 1
            If main.Cells(i, 7) = pc.Cells(3, 3) Then
                   
                acc_con = acc_con + 1
                acc_comp = acc_comp + 10
               
                gen_manag_con = gen_manag_con + 1
                gen_manag_comp = gen_manag_comp + 10
               
                cul_art_comp = cul_art_comp + 1
               
                mark_manag_comp = mark_manag_comp + 1
                   
            'Accounting 2
            ElseIf main.Cells(i, 7) = pc.Cells(4, 3) Then
           
                acc_con = acc_con + 1
                acc_comp = acc_comp + 10
               
                gen_manag_comp = gen_manag_comp + 1
           
            'Entrepreneurship
            ElseIf main.Cells(i, 7) = pc.Cells(5, 3) Or _
                   main.Cells(i, 7) = pc.Cells(6, 3) Then
           
                acc_comp = acc_comp + 1
               
                cul_art_comp = cul_art_comp + 1
               
                ece_comp = ece_comp + 1
               
                gen_manag_con = gen_manag_con + 1
                gen_manag_comp = gen_manag_comp + 10
               
                mark_manag_comp = mark_manag_comp + 1
               
                fam_con_sci_comp = fam_con_sci_comp + 1
               
                prog_soft_dev_comp = prog_soft_dev_comp + 1
           
            'Personal Finance
            ElseIf main.Cells(i, 7) = pc.Cells(7, 3) Or _
                   main.Cells(i, 7) = pc.Cells(8, 3) Then
           
                acc_comp = acc_comp + 1
               
                gen_manag_comp = gen_manag_comp + 1
               
                fam_con_sci_comp = fam_con_sci_comp + 1
             
            'Business Law
            ElseIf main.Cells(i, 7) = pc.Cells(9, 3) Then
           
                acc_comp = acc_comp + 1
               
                gen_manag_comp = gen_manag_comp + 1
               
                mark_manag_comp = mark_manag_comp + 1
           
            'Integrate Bus Ap1, Finance WB
            ElseIf main.Cells(i, 7) = pc.Cells(10, 3) Or _
                   main.Cells(i, 7) = pc.Cells(11, 3) Then
           
                acc_comp = acc_comp + 1
           
            'Agricultural Science and Technology
            ElseIf main.Cells(i, 7) = pc.Cells(3, 7) Then
               
                pl_an_sys_con = pl_an_sys_con + 10
                pl_an_sys_comp = pl_an_sys_comp + 10
               
                envi_nat_res_con = envi_nat_res_con + 1
                envi_nat_res_comp = envi_nat_res_comp + 10
               
                horti_comp = horti_comp + 10
           
            'Animal Science, Small Animal Care, Equine Science
            ElseIf main.Cells(i, 7) = pc.Cells(4, 7) Or _
                   main.Cells(i, 7) = pc.Cells(5, 7) Or _
                   main.Cells(i, 7) = pc.Cells(6, 7) Then
           
                pl_an_sys_con = pl_an_sys_con + 10
                pl_an_sys_comp = pl_an_sys_comp + 10
           
            'Intro to Veterinary Science H
            ElseIf main.Cells(i, 7) = pc.Cells(7, 7) Then
               
                pl_an_sys_con = pl_an_sys_con + 10
                pl_an_sys_comp = pl_an_sys_comp + 10
               
                health_sci_comp = health_sci_comp + 1
           
            'Ag work-based
            ElseIf main.Cells(i, 7) = pc.Cells(8, 7) Then
           
                pl_an_sys_comp = pl_an_sys_comp + 10
               
                envi_nat_res_comp = envi_nat_res_comp + 10
               
                horti_comp = horti_comp + 10
           
            'Auto Tech 1 - 2
            ElseIf main.Cells(i, 7) = pc.Cells(3, 11) Or _
                   main.Cells(i, 7) = pc.Cells(4, 11) Then
           
                auto_tech_con = auto_tech_con + 1
                auto_tech_comp = auto_tech_comp + 10
               
            'Auto Tech 3 - 4, WBL
            ElseIf main.Cells(i, 7) = pc.Cells(5, 11) Or _
                   main.Cells(i, 7) = pc.Cells(6, 11) Or _
                   main.Cells(i, 7) = pc.Cells(7, 11) Then
               
                auto_tech_comp = auto_tech_comp + 1
           
            'Culinary Arts Mgt 1
            ElseIf main.Cells(i, 7) = pc.Cells(3, 15) Then
           
                cul_art_con = cul_art_con + 1
                cul_art_comp = cul_art_comp + 10
               
                fam_con_sci_comp = fam_con_sci_comp + 1
           
            'Culinary Arts Mgt 2
            ElseIf main.Cells(i, 7) = pc.Cells(4, 15) Then
               
                cul_art_con = cul_art_con + 1
                cul_art_comp = cul_art_comp + 10
           
            'Food and Nutrition 1 - 2
            ElseIf main.Cells(i, 7) = pc.Cells(8, 15) Or _
                   main.Cells(i, 7) = pc.Cells(9, 15) Then
               
                cul_art_comp = cul_art_comp + 1
               
                ece_comp = ece_comp + 1
               
                fam_con_sci_con = fam_con_sci_con + 1
                fam_con_sci_comp = fam_con_sci_comp + 10
           
            'Web Page Design
            ElseIf main.Cells(i, 7) = pc.Cells(10, 15) Then
   
                cul_art_comp = cul_art_comp + 1
               
                gen_manag_comp = gen_manag_comp + 1
               
                prog_soft_dev_comp = prog_soft_dev_comp + 1
           
            'Digital Art and Design 1-2
            ElseIf main.Cells(i, 7) = pc.Cells(3, 19) Or _
                   main.Cells(i, 7) = pc.Cells(4, 19) Then
               
                dig_art_des_con = dig_art_des_con + 1
                dig_art_des_comp = dig_art_des_comp + 10
           
            'Digital Art and Design 3-4
            ElseIf main.Cells(i, 7) = pc.Cells(5, 19) Or _
                   main.Cells(i, 7) = pc.Cells(6, 19) Or _
                   main.Cells(i, 7) = pc.Cells(7, 19) Then
               
                dig_art_des_comp = dig_art_des_comp + 1
           
            'Art AV Tech & Comm Intern WBC
            ElseIf main.Cells(i, 7) = pc.Cells(8, 19) Then
           
                dig_art_des_comp = dig_art_des_comp + 1
               
                graph_comm_comp = graph_comm_comp + 1
               
                media_tech_comp = media_tech_comp + 1
           
            'Early Child Ed 1
            ElseIf main.Cells(i, 7) = pc.Cells(3, 23) Then
           
                ece_con = ece_con + 1
                ece_comp = ece_comp + 10
               
                fam_con_sci_con = fam_con_sci_con + 1
           
            'Early Childhood Ed 2
            ElseIf main.Cells(i, 7) = pc.Cells(4, 23) Then
           
                ece_con = ece_con + 1
                ece_comp = ece_comp + 10
           
            'Child Development 1
            ElseIf main.Cells(i, 7) = pc.Cells(5, 23) Then
               
                ece_comp = ece_comp + 10
               
                fam_con_sci_comp = fam_con_sci_comp + 10
       
            'Health Science 1
            ElseIf main.Cells(i, 7) = pc.Cells(10, 23) Or _
                   main.Cells(i, 7) = pc.Cells(11, 23) Then
           
                ece_comp = ece_comp + 1
               
                health_sci_con = health_sci_con + 1
                health_sci_comp = health_sci_comp + 10
               
                pltw_biomed_sci_comp = pltw_biomed_sci_comp + 1
               
                sport_med_comp = sport_med_comp + 1
           
            'Education & Training work base
            ElseIf main.Cells(i, 7) = pc.Cells(12, 23) Then
           
                ece_comp = ece_comp + 1
           
            'Teacher Cadet D
            ElseIf main.Cells(i, 7) = pc.Cells(13, 23) Then
           
                ece_comp = ece_comp + 1
               
                fam_con_sci_comp = fam_con_sci_comp + 1
           
            'DE Teacher Cadet Educational Psychology
            ElseIf main.Cells(i, 7) = pc.Cells(14, 23) Then
           
                ece_comp = ece_comp + 1
               
                fam_con_sci_comp = fam_con_sci_comp + 1
               
                prog_soft_dev = prog_soft_dev + 1
           
            'Env/Natural Res Management
            ElseIf main.Cells(i, 7) = pc.Cells(4, 27) Then
           
                envi_nat_res_con = envi_nat_res_con + 1
                envi_nat_res_comp = envi_nat_res_comp + 10
           
            'Wildlife Science, Outdoor Recreation
            ElseIf main.Cells(i, 7) = pc.Cells(5, 27) Or _
                   main.Cells(i, 7) = pc.Cells(6, 27) Then
           
                envi_nat_res_comp = envi_nat_res_comp + 10
           
            'Marketing, Marketing Management
            ElseIf main.Cells(i, 7) = pc.Cells(8, 31) Or _
                   main.Cells(i, 7) = pc.Cells(9, 31) Then
           
                gen_manag_comp = gen_manag_comp + 1
               
                mark_manag_con = mark_manag_con + 1
                mark_manag_comp = mark_manag_comp + 10
           
            'Graphic Communication 1 - 2
            ElseIf main.Cells(i, 7) = pc.Cells(3, 35) Or _
                   main.Cells(i, 7) = pc.Cells(4, 35) Then
           
                graph_comm_con = graph_comm_con + 1
                graph_comm_comp = graph_comm_comp + 10
           
            'Graphic Communication 3 - 4
            ElseIf main.Cells(i, 7) = pc.Cells(5, 35) Or _
                   main.Cells(i, 7) = pc.Cells(6, 35) Then
           
                graph_comm_comp = graph_comm_comp + 1
           
            'Health Science 2 CP
            ElseIf main.Cells(i, 7) = pc.Cells(5, 39) Then
           
                health_sci_con = health_sci_con + 1
                health_sci_comp = health_sci_comp + 10
               
                pltw_biomed_sci_comp = pltw_biomed_sci_comp + 1
               
                sport_med_comp = sport_med_comp + 1
           
            'Health Science 3 CP, Health Science 4 H Clinical Study
            ElseIf main.Cells(i, 7) = pc.Cells(6, 39) Or _
                   main.Cells(i, 7) = pc.Cells(8, 39) Then
           
                health_sci_comp = health_sci_comp + 1
               
                sport_med_comp = sport_med_comp + 1
           
            'Medical Terms
            ElseIf main.Cells(i, 7) = pc.Cells(7, 39) Then
           
                health_sci_comp = health_sci_comp + 1
               
            'Human Body Systems, Principles of Biomedical Science
            ElseIf main.Cells(i, 7) = pc.Cells(10, 39) Or _
                   main.Cells(i, 7) = pc.Cells(11, 39) Then
           
                health_sci_comp = health_sci_comp + 1
               
                pltw_biomed_sci_con = pltw_biomed_sci_con + 1
                pltw_biomed_sci_comp = pltw_biomed_sci_comp + 10
               
                sport_med_comp = sport_med_comp + 1
           
            'Sports Medicine 1 - 2
            ElseIf main.Cells(i, 7) = pc.Cells(12, 39) Or _
                   main.Cells(i, 7) = pc.Cells(13, 39) Then
           
                health_sci_comp = health_sci_comp + 1
               
                pltw_biomed_sci_comp = pltw_biomed_sci_comp + 1
               
                sport_med_comp = sport_med_comp + 10
                sport_med_con = sport_med_con + 1
           
            'Sports Medicine 3
            ElseIf main.Cells(i, 7) = pc.Cells(14, 39) Then
           
                health_sci_comp = health_sci_comp + 1
               
                pltw_biomed_sci_comp = pltw_biomed_sci_comp + 1
               
                sport_med_comp = sport_med_comp + 1
           
            'Intro to Horticulture
            ElseIf main.Cells(i, 7) = pc.Cells(3, 43) Then

                horti_con = horti_con + 1
                horti_comp = horti_comp + 10
           
            'Nursery Greenhouse and Garden Center Technology, Agribusiness and Marketing
            ElseIf main.Cells(i, 7) = pc.Cells(5, 43) Or _
                   main.Cells(i, 7) = pc.Cells(6, 43) Then
           
                horti_comp = horti_comp + 10
           
            'Sports & Entertainment Management
            ElseIf main.Cells(i, 7) = pc.Cells(5, 47) Or _
                   main.Cells(i, 7) = pc.Cells(6, 47) Then
               
                mark_manag_con = mark_manag_con + 1
                mark_manag_comp = mark_manag_comp + 10
           
            'Bus/M/CT WB
            ElseIf main.Cells(i, 7) = pc.Cells(11, 47) Then
           
                mark_manag_comp = mark_manag_comp + 1
           
            'Media Technology 1 - 2
            ElseIf main.Cells(i, 7) = pc.Cells(3, 51) Or _
                   main.Cells(i, 7) = pc.Cells(4, 51) Then
           
                media_tech_con = media_tech_con + 1
                media_tech_comp = media_tech_comp + 10
               
            'Media Technology 3 - 4
            ElseIf main.Cells(i, 7) = pc.Cells(5, 51) Or _
                   main.Cells(i, 7) = pc.Cells(6, 51) Or _
                   main.Cells(i, 7) = pc.Cells(7, 51) Or _
                   main.Cells(i, 7) = pc.Cells(8, 51) Then
           
                media_tech_comp = media_tech_comp + 1
           
            'Medical Interven
            ElseIf main.Cells(i, 7) = pc.Cells(5, 59) Then
           
                pltw_biomed_sci_comp = pltw_biomed_sci_comp + 1
           
            'Intro to Eng Design PLTW DC, Intro to Engineer, Prin/Engineering
            ElseIf main.Cells(i, 7) = pc.Cells(3, 63) Or _
                   main.Cells(i, 7) = pc.Cells(4, 63) Or _
                   main.Cells(i, 7) = pc.Cells(5, 63) Or _
                   main.Cells(i, 7) = pc.Cells(6, 63) Or _
                   main.Cells(i, 7) = pc.Cells(7, 63) Or _
                   main.Cells(i, 7) = pc.Cells(8, 63) Then
           
                pltw_pre_eng_con = pltw_pre_eng_con + 1
                pltw_pre_eng_comp = pltw_pre_eng_comp + 10
               
            'Aerospace Engineering, Civil Eng/Architec, Comp Integra Manu _
             Digital Electronic, Enginee Dsn&Dev, Environmental Sustainability
            ElseIf main.Cells(i, 7) = pc.Cells(10, 63) Or _
                   main.Cells(i, 7) = pc.Cells(11, 63) Or _
                   main.Cells(i, 7) = pc.Cells(12, 63) Or _
                   main.Cells(i, 7) = pc.Cells(13, 63) Or _
                   main.Cells(i, 7) = pc.Cells(14, 63) Or _
                   main.Cells(i, 7) = pc.Cells(15, 63) Or _
                   main.Cells(i, 7) = pc.Cells(16, 63) Then
           
                pltw_pre_eng_comp = pltw_pre_eng_comp + 1
           
            'Introduction to Computer Programming, Intermediate Computer Programming
            ElseIf main.Cells(i, 7) = pc.Cells(3, 75) Or _
                   main.Cells(i, 7) = pc.Cells(4, 75) Then
               
                prog_soft_dev_con = prog_soft_dev_con + 1
                prog_soft_dev_comp = prog_soft_dev_comp + 10
           
            'AP Computer Science Principles, Foundation in Animation, Fundamentals of Computing
            ElseIf main.Cells(i, 7) = pc.Cells(5, 75) Or _
                   main.Cells(i, 7) = pc.Cells(8, 75) Or _
                   main.Cells(i, 7) = pc.Cells(9, 75) Then
           
                prog_soft_dev_comp = prog_soft_dev_comp + 1
           
            End If
               
               
            'Check for concentrator status
            If acc_con = 2 Or _
               pl_an_sys_con = 50 Or _
               auto_tech_con = 2 Or _
               cul_art_con = 2 Or _
               dig_art_des_con = 2 Or _
               ece_con = 2 Or _
               envi_nat_res_con = 2 Or _
               gen_manag_con = 2 Or _
               graph_comm_con = 2 Or _
               health_sci_con = 2 Or _
               horti_con = 1 Or _
               mark_manag_con = 2 Or _
               media_tech_con = 2 Or _
               pltw_biomed_sci_con = 2 Or _
               pltw_pre_eng_con = 2 Or _
               pltw_pre_eng_con = 3 Or _
               sports_med_con = 2 Or _
               fam_con_sci_con = 2 Or _
               prog_soft_dev_con = 2 Then

                    pc.Activate
                    term_match = Application.Match(main.Cells(i, 14), pc.Range(Cells(4, 77), Cells(30, 77)), 0)
                   
                    If IsError(term_match) Then
                    ElseIf term_match > 0 Then
                   
                        If IsEmpty(pso.Cells(DestRow, 2)) Then
                           
                            pso.Cells(DestRow, 1) = main.Cells(i, 1)
                            pso.Cells(DestRow, 2) = "Y"
                            pso.Cells(DestRow, 6) = pc.Cells(term_match, 78)
                       
                        ElseIf Not IsEmpty(pso.Cells(DestRow, 2)) And IsEmpty(pso.Cells(DestRow, 8)) Then
                           
                            pso.Cells(DestRow, 1) = main.Cells(i, 1)
                            pso.Cells(DestRow, 8) = "Y"
                            pso.Cells(DestRow, 12) = pc.Cells(term_match, 78)
                       
                        End If
                       
                    End If
                   
                   
                    'reset concentrator
                    If acc_con = 2 Then
                        acc_con = 0
                    ElseIf pl_an_sys_con = 50 Then
                        pl_an_sys_con = 0
                    ElseIf auto_tech_con = 2 Then
                        auto_tech_con = 0
                    ElseIf cul_art_con = 2 Then
                        cul_art_con = 0
                    ElseIf dig_art_des_con = 2 Then
                        dig_art_des_con = 0
                    ElseIf ece_con = 2 Then
                        ece_con = 0
                    ElseIf envi_nat_res_con = 2 Then
                        envi_nat_res_con = 0
                    ElseIf gen_manag_con = 2 Then
                        gen_manag_con = 0
                    ElseIf graph_comm_con = 2 Then
                        graph_comm_con = 0
                    ElseIf health_sci_con = 2 Then
                        health_sci_con = 0
                    ElseIf horti_con = 1 Then
                        horti_con = 0
                    ElseIf mark_manag_con = 2 Then
                        mark_manag_con = 0
                    ElseIf media_tech_con = 2 Then
                        media_tech_con = 0
                    ElseIf pltw_biomed_sci_con = 2 Then
                        pltw_biomed_sci_con = 0
                    ElseIf pltw_pre_eng_con = 2 Then
                        pltw_pre_eng_con = 0
                    ElseIf pltw_pre_eng_con = 3 Then
                        pltw_pre_eng_con = 0
                    ElseIf sports_med_con = 2 Then
                        sports_med_con = 0
                    ElseIf fam_con_sci_con = 2 Then
                        fam_con_sci_con = 0
                    ElseIf prog_soft_dev_con = 2 Then
                        prog_soft_dev_con = 0
                    End If
               
            'Check for completer status
            ElseIf (acc_comp >= 21 And acc_comp <= 25) Or _
                   pl_an_sys_comp = 50 Or _
                   pl_an_sys_comp = 60 Or _
                   (auto_tech_comp >= 21 And auto_tech_comp <= 23) Or _
                   (cul_art_comp >= 21 And cul_art_comp <= 25) Or _
                   (dig_art_des_comp >= 21 And dig_art_des_comp <= 23) Or _
                   (ece_comp >= 31 And ece_comp <= 37) Or _
                   envi_nat_res_comp = 40 Or _
                   envi_nat_res_comp = 50 Or _
                   (gen_manag_comp >= 21 And gen_manag_comp <= 26) Or _
                   (graph_comm_comp >= 21 And graph_comm_comp <= 23) Or _
                   (health_sci_comp >= 21 And health_sci_comp <= 28) Or _
                   horti_comp = 40 Or _
                   horti_comp = 50 Or _
                   (mark_manag_comp >= 21 And mark_manag_comp <= 24) Or _
                   mark_manag_comp = 30 Or _
                   (media_tech_comp >= 21 And media_tech_comp <= 23) Or _
                   (pltw_biomed_sci_comp >= 21 And pltw_biomed_sci_comp <= 26) Or _
                   (pltw_pre_eng_comp >= 22 And pltw_pre_eng_comp <= 26) Or _
                   (sports_med_comp >= 21 And sports_med_comp <= 27) Or _
                   (fam_con_sci_comp >= 21 And fam_con_sci_comp <= 26) Or _
                   fam_con_sci_comp = 30 Or _
                   (prog_soft_dev_comp >= 21 And prog_soft_dev_comp <= 25) Then
                   
                        pc.Activate
                        term_match = Application.Match(main.Cells(i, 14), pc.Range(Cells(4, 77), Cells(30, 77)), 0)
                       
                        If IsError(term_match) Then
                        ElseIf term_match > 0 Then
                       
                            If IsEmpty(pso.Cells(DestRow, 3)) Then
                           
                                pso.Cells(DestRow, 1) = main.Cells(i, 1)
                                pso.Cells(DestRow, 3) = "Y"
                                pso.Cells(DestRow, 5) = pc.Cells(term_match, 78)
                           
                            ElseIf Not IsEmpty(pso.Cells(DestRow, 3)) And IsEmpty(pso.Cells(DestRow, 9)) Then
                           
                                pso.Cells(DestRow, 1) = main.Cells(i, 1)
                                pso.Cells(DestRow, 9) = "Y"
                                pso.Cells(DestRow, 11) = pc.Cells(term_match, 78)
                               
                            End If
                       
                        End If
                   
                       
                        'reset completer
                        If acc_comp >= 21 And acc_comp <= 25 Then
                            acc_comp = 0
                        ElseIf pl_an_sys_comp = 50 Or _
                               pl_an_sys_comp = 60 Then
                                    pl_an_sys_comp = 0
                        ElseIf auto_tech_comp >= 21 And auto_tech_comp <= 23 Then
                            auto_tech_comp = 0
                        ElseIf cul_art_comp >= 21 And cul_art_comp <= 25 Then
                            cul_art_comp = 0
                        ElseIf dig_art_des_comp >= 21 And dig_art_des_comp <= 23 Then
                            dig_art_des_comp = 0
                        ElseIf ece_comp >= 31 And ece_comp <= 37 Then
                            ece_comp = 0
                        ElseIf envi_nat_res_comp = 40 Or _
                               envi_nat_res_comp = 50 Then
                                    envi_nat_res_comp = 0
                        ElseIf gen_manag_comp >= 21 And gen_manag_comp <= 26 Then
                            gen_manag_comp = 0
                        ElseIf graph_comm_comp >= 21 Or graph_comm_comp <= 23 Then
                            graph_comm_comp = 0
                        ElseIf health_sci_comp >= 21 Or health_sci_comp <= 28 Then
                            health_sci_comp = 0
                        ElseIf horti_comp = 40 Or _
                               horti_comp = 50 Then
                                    horti_comp = 0
                        ElseIf mark_manag_comp >= 21 And mark_manag_comp <= 24 Or _
                               mark_manag_comp = 30 Then
                                    mark_manag_comp = 0
                        ElseIf media_tech_comp >= 21 And media_tech_comp <= 23 Then
                            media_tech_comp = 0
                        ElseIf pltw_biomed_sci_comp >= 21 And pltw_biomed_sci_comp <= 26 Then
                            pltw_biomed_sci_comp = 0
                        ElseIf pltw_pre_eng_comp >= 22 And pltw_pre_eng_comp <= 26 Then
                            pltw_pre_eng_comp = 0
                        ElseIf sports_med_comp >= 21 And sports_med_comp <= 27 Then
                            sports_med_comp = 0
                        ElseIf fam_con_sci_comp >= 21 And fam_con_sci_comp <= 26 Or _
                               fam_con_sci_comp = 30 Then
                                    fam_con_sci_comp = 0
                        ElseIf prog_soft_dev_comp >= 21 And prog_soft_dev_comp <= 25 Then
                            prog_soft_dev_comp = 0
                        End If
                       
            End If
                           
                           
         Else
                   
            'Accounting 1
            If main.Cells(i, 7) = pc.Cells(3, 3) Then
                   
                acc_con = acc_con + 1
                acc_comp = acc_comp + 10
               
                gen_manag_con = gen_manag_con + 1
                gen_manag_comp = gen_manag_comp + 10
               
                cul_art_comp = cul_art_comp + 1
               
                mark_manag_comp = mark_manag_comp + 1
                   
            'Accounting 2
            ElseIf main.Cells(i, 7) = pc.Cells(4, 3) Then
           
                acc_con = acc_con + 1
                acc_comp = acc_comp + 10
               
                gen_manag_comp = gen_manag_comp + 1
           
            'Entrepreneurship
            ElseIf main.Cells(i, 7) = pc.Cells(5, 3) Or _
                   main.Cells(i, 7) = pc.Cells(6, 3) Then
           
                acc_comp = acc_comp + 1
               
                cul_art_comp = cul_art_comp + 1
               
                ece_comp = ece_comp + 1
               
                gen_manag_con = gen_manag_con + 1
                gen_manag_comp = gen_manag_comp + 10
               
                mark_manag_comp = mark_manag_comp + 1
               
                fam_con_sci_comp = fam_con_sci_comp + 1
               
                prog_soft_dev_comp = prog_soft_dev_comp + 1
           
            'Personal Finance
            ElseIf main.Cells(i, 7) = pc.Cells(7, 3) Or _
                   main.Cells(i, 7) = pc.Cells(8, 3) Then
           
                acc_comp = acc_comp + 1
               
                gen_manag_comp = gen_manag_comp + 1
               
                fam_con_sci_comp = fam_con_sci_comp + 1
             
            'Business Law
            ElseIf main.Cells(i, 7) = pc.Cells(9, 3) Then
           
                acc_comp = acc_comp + 1
               
                gen_manag_comp = gen_manag_comp + 1
               
                mark_manag_comp = mark_manag_comp + 1
           
            'Integrate Bus Ap1, Finance WB
            ElseIf main.Cells(i, 7) = pc.Cells(10, 3) Or _
                   main.Cells(i, 7) = pc.Cells(11, 3) Then
           
                acc_comp = acc_comp + 1
           
            'Agricultural Science and Technology
            ElseIf main.Cells(i, 7) = pc.Cells(3, 7) Then
               
                pl_an_sys_con = pl_an_sys_con + 10
                pl_an_sys_comp = pl_an_sys_comp + 10
               
                envi_nat_res_con = envi_nat_res_con + 1
                envi_nat_res_comp = envi_nat_res_comp + 10
               
                horti_comp = horti_comp + 10
           
            'Animal Science, Small Animal Care, Equine Science
            ElseIf main.Cells(i, 7) = pc.Cells(4, 7) Or _
                   main.Cells(i, 7) = pc.Cells(5, 7) Or _
                   main.Cells(i, 7) = pc.Cells(6, 7) Then
           
                pl_an_sys_con = pl_an_sys_con + 10
                pl_an_sys_comp = pl_an_sys_comp + 10
           
            'Intro to Veterinary Science H
            ElseIf main.Cells(i, 7) = pc.Cells(7, 7) Then
               
                pl_an_sys_con = pl_an_sys_con + 10
                pl_an_sys_comp = pl_an_sys_comp + 10
               
                health_sci_comp = health_sci_comp + 1
           
            'Ag work-based
            ElseIf main.Cells(i, 7) = pc.Cells(8, 7) Then
           
                pl_an_sys_comp = pl_an_sys_comp + 10
               
                envi_nat_res_comp = envi_nat_res_comp + 10
               
                horti_comp = horti_comp + 10
           
            'Auto Tech 1 - 2
            ElseIf main.Cells(i, 7) = pc.Cells(3, 11) Or _
                   main.Cells(i, 7) = pc.Cells(4, 11) Then
           
                auto_tech_con = auto_tech_con + 1
                auto_tech_comp = auto_tech_comp + 10
               
            'Auto Tech 3 - 4, WBL
            ElseIf main.Cells(i, 7) = pc.Cells(5, 11) Or _
                   main.Cells(i, 7) = pc.Cells(6, 11) Or _
                   main.Cells(i, 7) = pc.Cells(7, 11) Then
               
                auto_tech_comp = auto_tech_comp + 1
           
            'Culinary Arts Mgt 1
            ElseIf main.Cells(i, 7) = pc.Cells(3, 15) Then
           
                cul_art_con = cul_art_con + 1
                cul_art_comp = cul_art_comp + 10
               
                fam_con_sci_comp = fam_con_sci_comp + 1
           
            'Culinary Arts Mgt 2
            ElseIf main.Cells(i, 7) = pc.Cells(4, 15) Then
               
                cul_art_con = cul_art_con + 1
                cul_art_comp = cul_art_comp + 10
           
            'Food and Nutrition 1 - 2
            ElseIf main.Cells(i, 7) = pc.Cells(8, 15) Or _
                   main.Cells(i, 7) = pc.Cells(9, 15) Then
               
                cul_art_comp = cul_art_comp + 1
               
                ece_comp = ece_comp + 1
               
                fam_con_sci_con = fam_con_sci_con + 1
                fam_con_sci_comp = fam_con_sci_comp + 10
           
            'Web Page Design
            ElseIf main.Cells(i, 7) = pc.Cells(10, 15) Then
   
                cul_art_comp = cul_art_comp + 1
               
                gen_manag_comp = gen_manag_comp + 1
               
                prog_soft_dev_comp = prog_soft_dev_comp + 1
           
            'Digital Art and Design 1-2
            ElseIf main.Cells(i, 7) = pc.Cells(3, 19) Or _
                   main.Cells(i, 7) = pc.Cells(4, 19) Then
               
                dig_art_des_con = dig_art_des_con + 1
                dig_art_des_comp = dig_art_des_comp + 10
           
            'Digital Art and Design 3-4
            ElseIf main.Cells(i, 7) = pc.Cells(5, 19) Or _
                   main.Cells(i, 7) = pc.Cells(6, 19) Or _
                   main.Cells(i, 7) = pc.Cells(7, 19) Then
               
                dig_art_des_comp = dig_art_des_comp + 1
           
            'Art AV Tech & Comm Intern WBC
            ElseIf main.Cells(i, 7) = pc.Cells(8, 19) Then
           
                dig_art_des_comp = dig_art_des_comp + 1
               
                graph_comm_comp = graph_comm_comp + 1
               
                media_tech_comp = media_tech_comp + 1
           
            'Early Child Ed 1
            ElseIf main.Cells(i, 7) = pc.Cells(3, 23) Then
           
                ece_con = ece_con + 1
                ece_comp = ece_comp + 10
               
                fam_con_sci_con = fam_con_sci_con + 1
           
            'Early Childhood Ed 2
            ElseIf main.Cells(i, 7) = pc.Cells(4, 23) Then
           
                ece_con = ece_con + 1
                ece_comp = ece_comp + 10
           
            'Child Development 1
            ElseIf main.Cells(i, 7) = pc.Cells(5, 23) Then
               
                ece_comp = ece_comp + 10
               
                fam_con_sci_comp = fam_con_sci_comp + 10
       
            'Health Science 1
            ElseIf main.Cells(i, 7) = pc.Cells(10, 23) Or _
                   main.Cells(i, 7) = pc.Cells(11, 23) Then
           
                ece_comp = ece_comp + 1
               
                health_sci_con = health_sci_con + 1
                health_sci_comp = health_sci_comp + 10
               
                pltw_biomed_sci_comp = pltw_biomed_sci_comp + 1
               
                sport_med_comp = sport_med_comp + 1
           
            'Education & Training work base
            ElseIf main.Cells(i, 7) = pc.Cells(12, 23) Then
           
                ece_comp = ece_comp + 1
           
            'Teacher Cadet D
            ElseIf main.Cells(i, 7) = pc.Cells(13, 23) Then
           
                ece_comp = ece_comp + 1
               
                fam_con_sci_comp = fam_con_sci_comp + 1
           
            'DE Teacher Cadet Educational Psychology
            ElseIf main.Cells(i, 7) = pc.Cells(14, 23) Then
           
                ece_comp = ece_comp + 1
               
                fam_con_sci_comp = fam_con_sci_comp + 1
               
                prog_soft_dev = prog_soft_dev + 1
           
            'Env/Natural Res Management
            ElseIf main.Cells(i, 7) = pc.Cells(4, 27) Then
           
                envi_nat_res_con = envi_nat_res_con + 1
                envi_nat_res_comp = envi_nat_res_comp + 10
           
            'Wildlife Science, Outdoor Recreation
            ElseIf main.Cells(i, 7) = pc.Cells(5, 27) Or _
                   main.Cells(i, 7) = pc.Cells(6, 27) Then
           
                envi_nat_res_comp = envi_nat_res_comp + 10
           
            'Marketing, Marketing Management
            ElseIf main.Cells(i, 7) = pc.Cells(8, 31) Or _
                   main.Cells(i, 7) = pc.Cells(9, 31) Then
           
                gen_manag_comp = gen_manag_comp + 1
               
                mark_manag_con = mark_manag_con + 1
                mark_manag_comp = mark_manag_comp + 10
           
            'Graphic Communication 1 - 2
            ElseIf main.Cells(i, 7) = pc.Cells(3, 35) Or _
                   main.Cells(i, 7) = pc.Cells(4, 35) Then
           
                graph_comm_con = graph_comm_con + 1
                graph_comm_comp = graph_comm_comp + 10
           
            'Graphic Communication 3 - 4
            ElseIf main.Cells(i, 7) = pc.Cells(5, 35) Or _
                   main.Cells(i, 7) = pc.Cells(6, 35) Then
           
                graph_comm_comp = graph_comm_comp + 1
           
            'Health Science 2 CP
            ElseIf main.Cells(i, 7) = pc.Cells(5, 39) Then
           
                health_sci_con = health_sci_con + 1
                health_sci_comp = health_sci_comp + 10
               
                pltw_biomed_sci_comp = pltw_biomed_sci_comp + 1
               
                sport_med_comp = sport_med_comp + 1
           
            'Health Science 3 CP, Health Science 4 H Clinical Study
            ElseIf main.Cells(i, 7) = pc.Cells(6, 39) Or _
                   main.Cells(i, 7) = pc.Cells(8, 39) Then
           
                health_sci_comp = health_sci_comp + 1
               
                sport_med_comp = sport_med_comp + 1
           
            'Medical Terms
            ElseIf main.Cells(i, 7) = pc.Cells(7, 39) Then
           
                health_sci_comp = health_sci_comp + 1
               
            'Human Body Systems, Principles of Biomedical Science
            ElseIf main.Cells(i, 7) = pc.Cells(10, 39) Or _
                   main.Cells(i, 7) = pc.Cells(11, 39) Then
           
                health_sci_comp = health_sci_comp + 1
               
                pltw_biomed_sci_con = pltw_biomed_sci_con + 1
                pltw_biomed_sci_comp = pltw_biomed_sci_comp + 10
               
                sport_med_comp = sport_med_comp + 1
           
            'Sports Medicine 1 - 2
            ElseIf main.Cells(i, 7) = pc.Cells(12, 39) Or _
                   main.Cells(i, 7) = pc.Cells(13, 39) Then
           
                health_sci_comp = health_sci_comp + 1
               
                pltw_biomed_sci_comp = pltw_biomed_sci_comp + 1
               
                sport_med_comp = sport_med_comp + 10
                sport_med_con = sport_med_con + 1
           
            'Sports Medicine 3
            ElseIf main.Cells(i, 7) = pc.Cells(14, 39) Then
           
                health_sci_comp = health_sci_comp + 1
               
                pltw_biomed_sci_comp = pltw_biomed_sci_comp + 1
               
                sport_med_comp = sport_med_comp + 1
           
            'Intro to Horticulture
            ElseIf main.Cells(i, 7) = pc.Cells(3, 43) Then

                horti_con = horti_con + 1
                horti_comp = horti_comp + 10
           
            'Nursery Greenhouse and Garden Center Technology, Agribusiness and Marketing
            ElseIf main.Cells(i, 7) = pc.Cells(5, 43) Or _
                   main.Cells(i, 7) = pc.Cells(6, 43) Then
           
                horti_comp = horti_comp + 10
           
            'Sports & Entertainment Management
            ElseIf main.Cells(i, 7) = pc.Cells(5, 47) Or _
                   main.Cells(i, 7) = pc.Cells(6, 47) Then
               
                mark_manag_con = mark_manag_con + 1
                mark_manag_comp = mark_manag_comp + 10
           
            'Bus/M/CT WB
            ElseIf main.Cells(i, 7) = pc.Cells(11, 47) Then
           
                mark_manag_comp = mark_manag_comp + 1
           
            'Media Technology 1 - 2
            ElseIf main.Cells(i, 7) = pc.Cells(3, 51) Or _
                   main.Cells(i, 7) = pc.Cells(4, 51) Then
           
                media_tech_con = media_tech_con + 1
                media_tech_comp = media_tech_comp + 10
               
            'Media Technology 3 - 4
            ElseIf main.Cells(i, 7) = pc.Cells(5, 51) Or _
                   main.Cells(i, 7) = pc.Cells(6, 51) Or _
                   main.Cells(i, 7) = pc.Cells(7, 51) Or _
                   main.Cells(i, 7) = pc.Cells(8, 51) Then
           
                media_tech_comp = media_tech_comp + 1
           
            'Medical Interven
            ElseIf main.Cells(i, 7) = pc.Cells(5, 59) Then
           
                pltw_biomed_sci_comp = pltw_biomed_sci_comp + 1
           
            'Intro to Eng Design PLTW DC, Intro to Engineer, Prin/Engineering
            ElseIf main.Cells(i, 7) = pc.Cells(3, 63) Or _
                   main.Cells(i, 7) = pc.Cells(4, 63) Or _
                   main.Cells(i, 7) = pc.Cells(5, 63) Or _
                   main.Cells(i, 7) = pc.Cells(6, 63) Or _
                   main.Cells(i, 7) = pc.Cells(7, 63) Or _
                   main.Cells(i, 7) = pc.Cells(8, 63) Then
           
                pltw_pre_eng_con = pltw_pre_eng_con + 1
                pltw_pre_eng_comp = pltw_pre_eng_comp + 10
               
            'Aerospace Engineering, Civil Eng/Architec, Comp Integra Manu _
             Digital Electronic, Enginee Dsn&Dev, Environmental Sustainability
            ElseIf main.Cells(i, 7) = pc.Cells(10, 63) Or _
                   main.Cells(i, 7) = pc.Cells(11, 63) Or _
                   main.Cells(i, 7) = pc.Cells(12, 63) Or _
                   main.Cells(i, 7) = pc.Cells(13, 63) Or _
                   main.Cells(i, 7) = pc.Cells(14, 63) Or _
                   main.Cells(i, 7) = pc.Cells(15, 63) Or _
                   main.Cells(i, 7) = pc.Cells(16, 63) Then
           
                pltw_pre_eng_comp = pltw_pre_eng_comp + 1
           
            'Introduction to Computer Programming, Intermediate Computer Programming
            ElseIf main.Cells(i, 7) = pc.Cells(3, 75) Or _
                   main.Cells(i, 7) = pc.Cells(4, 75) Then
               
                prog_soft_dev_con = prog_soft_dev_con + 1
                prog_soft_dev_comp = prog_soft_dev_comp + 10
           
            'AP Computer Science Principles, Foundation in Animation, Fundamentals of Computing
            ElseIf main.Cells(i, 7) = pc.Cells(5, 75) Or _
                   main.Cells(i, 7) = pc.Cells(8, 75) Or _
                   main.Cells(i, 7) = pc.Cells(9, 75) Then
           
                prog_soft_dev_comp = prog_soft_dev_comp + 1
           
            End If
               
               
            'Check for concentrator status
            If acc_con = 2 Or _
               pl_an_sys_con = 50 Or _
               auto_tech_con = 2 Or _
               cul_art_con = 2 Or _
               dig_art_des_con = 2 Or _
               ece_con = 2 Or _
               envi_nat_res_con = 2 Or _
               gen_manag_con = 2 Or _
               graph_comm_con = 2 Or _
               health_sci_con = 2 Or _
               horti_con = 1 Or _
               mark_manag_con = 2 Or _
               media_tech_con = 2 Or _
               pltw_biomed_sci_con = 2 Or _
               pltw_pre_eng_con = 2 Or _
               pltw_pre_eng_con = 3 Or _
               sports_med_con = 2 Or _
               fam_con_sci_con = 2 Or _
               prog_soft_dev_con = 2 Then

                    pc.Activate
                    term_match = Application.Match(main.Cells(i, 14), pc.Range(Cells(4, 77), Cells(30, 77)), 0)
                   
                    If IsError(term_match) Then
                    ElseIf term_match > 0 Then
                   
                        If IsEmpty(pso.Cells(DestRow, 2)) Then
                           
                            pso.Cells(DestRow, 1) = main.Cells(i, 1)
                            pso.Cells(DestRow, 2) = "Y"
                            pso.Cells(DestRow, 6) = pc.Cells(term_match, 78)
                       
                        ElseIf Not IsEmpty(pso.Cells(DestRow, 2)) And IsEmpty(pso.Cells(DestRow, 8)) Then
                           
                            pso.Cells(DestRow, 1) = main.Cells(i, 1)
                            pso.Cells(DestRow, 8) = "Y"
                            pso.Cells(DestRow, 12) = pc.Cells(term_match, 78)
                       
                        Else
                        End If
                   
                    End If
                   
                   
                   
               
            'Check for completer status
            ElseIf (acc_comp >= 21 And acc_comp <= 25) Or _
                   pl_an_sys_comp = 50 Or _
                   pl_an_sys_comp = 60 Or _
                   (auto_tech_comp >= 21 And auto_tech_comp <= 23) Or _
                   (cul_art_comp >= 21 And cul_art_comp <= 25) Or _
                   (dig_art_des_comp >= 21 And dig_art_des_comp <= 23) Or _
                   (ece_comp >= 31 And ece_comp <= 37) Or _
                   envi_nat_res_comp = 40 Or _
                   envi_nat_res_comp = 50 Or _
                   (gen_manag_comp >= 21 And gen_manag_comp <= 26) Or _
                   (graph_comm_comp >= 21 And graph_comm_comp <= 23) Or _
                   (health_sci_comp >= 21 And health_sci_comp <= 28) Or _
                   horti_comp = 40 Or _
                   horti_comp = 50 Or _
                   (mark_manag_comp >= 21 And mark_manag_comp <= 24) Or _
                   mark_manag_comp = 30 Or _
                   (media_tech_comp >= 21 And media_tech_comp <= 23) Or _
                   (pltw_biomed_sci_comp >= 21 And pltw_biomed_sci_comp <= 26) Or _
                   (pltw_pre_eng_comp >= 22 And pltw_pre_eng_comp <= 26) Or _
                   (sports_med_comp >= 21 And sports_med_comp <= 27) Or _
                   (fam_con_sci_comp >= 21 And fam_con_sci_comp <= 26) Or _
                   fam_con_sci_comp = 30 Or _
                   (prog_soft_dev_comp >= 21 And prog_soft_dev_comp <= 25) Then
                   
                        pc.Activate
                        term_match = Application.Match(main.Cells(i, 14), pc.Range(Cells(4, 77), Cells(30, 77)), 0)
                       
                        If IsError(term_match) Then
                        ElseIf term_match > 0 Then
                       
                            If IsEmpty(pso.Cells(DestRow, 3)) Then
                           
                                pso.Cells(DestRow, 1) = main.Cells(i, 1)
                                pso.Cells(DestRow, 3) = "Y"
                                pso.Cells(DestRow, 5) = pc.Cells(term_match, 78)
                           
                            ElseIf Not IsEmpty(pso.Cells(DestRow, 3)) And IsEmpty(pso.Cells(DestRow, 9)) Then
                           
                                pso.Cells(DestRow, 1) = main.Cells(i, 1)
                                pso.Cells(DestRow, 9) = "Y"
                                pso.Cells(DestRow, 11) = pc.Cells(term_match, 78)
                               
                            Else
                            End If
                       
                        End If
           
            End If
           
            'reset all concentrator and completer counts (next line is a new student)
            acc_con = 0
            acc_comp = 0
            pl_an_sys_con = 0
            pl_an_sys_comp = 0
            auto_tech_con = 0
            auto_tech_comp = 0
            cul_art_con = 0
            cul_art_comp = 0
            dig_art_des_con = 0
            dig_art_des_comp = 0
            ece_con = 0
            ece_comp = 0
            envi_nat_res_con = 0
            envi_nat_res_comp = 0
            gen_manag_con = 0
            gen_manag_comp = 0
            graph_comm_con = 0
            graph_comm_comp = 0
            health_sci_con = 0
            health_sci_comp = 0
            horti_con = 0
            horti_comp = 0
            mark_manag_con = 0
            mark_manag_comp = 0
            media_tech_con = 0
            media_tech_comp = 0
            pltw_biomed_sci_con = 0
            pltw_biomed_sci_comp = 0
            pltw_pre_eng_con = 0
            pltw_pre_eng_comp = 0
            sport_med_con = 0
            sport_med_comp = 0
            fam_con_sci_con = 0
            fam_con_sci_comp = 0
            prog_soft_dev_con = 0
            prog_soft_dev_comp = 0
           
            DestRow = pso.Range("A" & main.Rows.Count).End(xlUp).Offset(1, 0).Row
       
        End If

    Next i

End Sub

```
