library(openxlsx)
library(tidyverse)
library(dplyr)
library(redcapAPI)
library(RMySQL)
library(summarytools)
library(readxl)
library(haven)
library(xlsx)
library(survival)
library(conflicted)

wb1 <- xlsx::loadWorkbook("Metadata/TB Omni Report Template.xlsx")

works_sheets <- xlsx::getSheets(wb1)

tmp_sheet <- works_sheets[["Data"]]

rows <- getRows(tmp_sheet)

cells <- getCells(rows)

today <- format(Sys.time(), "%Y-%m-%d")

filename_new <- paste("Data/TB Omni Report ",today,".xlsx")



############################################Index Enrolment
setCellValue(cells[["3.2"]], nrow(subset(raw_data_index_enrolment_arm_1, !is.na(raw_data_index_enrolment_arm_1$approach_intro))))

setCellValue(cells[["4.2"]], nrow(subset(raw_data_index_enrolment_arm_1, raw_data_index_enrolment_arm_1$s_q4=='Proceed')))


setCellValue(cells[["6.2"]], nrow(subset(raw_data_index_enrolment_arm_1, raw_data_index_enrolment_arm_1$s_q13=='No')))

setCellValue(cells[["7.2"]], nrow(subset(raw_data_index_enrolment_arm_1, raw_data_index_enrolment_arm_1$s_q15=='No')))

setCellValue(cells[["8.2"]], nrow(subset(raw_data_index_enrolment_arm_1, raw_data_index_enrolment_arm_1$s_q8=='No')))

setCellValue(cells[["9.2"]], nrow(subset(raw_data_index_enrolment_arm_1, raw_data_index_enrolment_arm_1$s_q11 <18)))

setCellValue(cells[["10.2"]], nrow(subset(raw_data_index_enrolment_arm_1, raw_data_index_enrolment_arm_1$bcm_yn=='No')))


#Eligible
setCellValue(cells[["11.2"]], nrow(subset(raw_data_index_enrolment_arm_1, raw_data_index_enrolment_arm_1$s_q16=='Proceed')))

setCellValue(cells[["12.2"]], nrow(subset(raw_data_index_enrolment_arm_1, raw_data_index_enrolment_arm_1$tb_tf_study_time_point=='HHCI')))

setCellValue(cells[["13.2"]], nrow(subset(raw_data_index_enrolment_arm_1, raw_data_index_enrolment_arm_1$s_q20=='Yes')))


################################################HH Visiting
setCellValue(cells[["18.2"]], nrow(subset(raw_data_index_enrolment_arm_1, raw_data_index_enrolment_arm_1$s_q20=='Yes')))

setCellValue(cells[["19.2"]], nrow(subset(raw_data_index_hhc_investig_arm_1, !is.na(raw_data_index_hhc_investig_arm_1$name_hhc))))

setCellValue(cells[["20.2"]], nrow(raw_data_index_hhc_investig_arm_1%>% dplyr::filter(!is.na(visit1_outcome)) %>% distinct(record_id)))

setCellValue(cells[["23.2"]], nrow(subset(raw_data_index_hhc_investig_arm_1, raw_data_index_hhc_investig_arm_1$visit1_outcome=='Household member present' |
                                     raw_data_index_hhc_investig_arm_1$visit2_outcomes=='Household member present' |
                                     raw_data_index_hhc_investig_arm_1$visit3_outcomes=='Household member present')))

setCellValue(cells[["24.2"]], nrow(subset(raw_data_index_hhc_investig_arm_1, raw_data_index_hhc_investig_arm_1$visit1_outcome=='Refused investigation' |
                                     raw_data_index_hhc_investig_arm_1$visit2_outcomes=='Refused investigation' |
                                     raw_data_index_hhc_investig_arm_1$visit3_outcomes=='Refused investigation')))

setCellValue(cells[["25.2"]], nrow(subset(raw_data_index_hhc_investig_arm_1, raw_data_index_hhc_investig_arm_1$visit1_outcome=='Contacts does not reside in this household' |
                                     raw_data_index_hhc_investig_arm_1$visit2_outcomes=='Contacts does not reside in this household' |
                                     raw_data_index_hhc_investig_arm_1$visit3_outcomes=='Contacts does not reside in this household')))

setCellValue(cells[["26.2"]], nrow(subset(raw_data_index_hhc_investig_arm_1, raw_data_index_hhc_investig_arm_1$visit1_outcome=='Household contact relocated' |
                                     raw_data_index_hhc_investig_arm_1$visit2_outcomes=='Household contact relocated' |
                                     raw_data_index_hhc_investig_arm_1$visit3_outcomes=='Household contact relocated')))


setCellValue(cells[["27.2"]], nrow(subset(raw_data_index_hhc_investig_arm_1, raw_data_index_hhc_investig_arm_1$visit1_outcome=='Household not found' |
                                     raw_data_index_hhc_investig_arm_1$visit2_outcomes=='Household not found' |
                                     raw_data_index_hhc_investig_arm_1$visit3_outcomes=='Household not found')))


setCellValue(cells[["28.2"]], nrow(subset(raw_data_index_hhc_investig_arm_1, raw_data_index_hhc_investig_arm_1$visit1_outcome=='Household member not present' &
                                     (raw_data_index_hhc_investig_arm_1$visit2_outcomes=='Household member not present') &
                                     raw_data_index_hhc_investig_arm_1$visit3_outcomes=='Household member not present')))

setCellValue(cells[["29.2"]], nrow(subset(raw_data_index_hhc_investig_arm_1, raw_data_index_hhc_investig_arm_1$visit1_outcome=='Household found but only index patient present and they do not have HHCs' |
                                     raw_data_index_hhc_investig_arm_1$visit2_outcomes=='Household found but only index patient present and they do not have HHCs' |
                                     raw_data_index_hhc_investig_arm_1$visit3_outcomes=='Household found but only index patient present and they do not have HHCs')))


setCellValue(cells[["30.2"]], nrow(subset(raw_data_index_hhc_investig_arm_1, raw_data_index_hhc_investig_arm_1$visit1_outcome=='Household found and knows index but index stays elsewhere' |
                                     raw_data_index_hhc_investig_arm_1$visit2_outcomes=='Household found and knows index but index stays elsewhere' |
                                     raw_data_index_hhc_investig_arm_1$visit3_outcomes=='Household found and knows index but index stays elsewhere')))


setCellValue(cells[["31.2"]], nrow(subset(raw_data_index_hhc_investig_arm_1, raw_data_index_hhc_investig_arm_1$visit1_outcome=='Household found but claims no relation to index patient' |
                                     raw_data_index_hhc_investig_arm_1$visit2_outcomes=='Household found but claims no relation to index patient' |
                                     raw_data_index_hhc_investig_arm_1$visit3_outcomes=='Household found but claims no relation to index patient')))


setCellValue(cells[["32.2"]], nrow(subset(raw_data_index_hhc_investig_arm_1, raw_data_index_hhc_investig_arm_1$vc_hhm_sc_sc=='Yes')))

setCellValue(cells[["34.2"]], nrow(subset(raw_data_index_hhc_investig_arm_1, raw_data_index_hhc_investig_arm_1$hhm_age_sc_calc<18)))

setCellValue(cells[["35.2"]], nrow(subset(raw_data_index_hhc_investig_arm_1, raw_data_index_hhc_investig_arm_1$tb_past_treat_sc=='Yes')))

setCellValue(cells[["36.2"]], nrow(subset(raw_data_index_hhc_investig_arm_1, raw_data_index_hhc_investig_arm_1$vc_hhm_sc_sc=='No')))

setCellValue(cells[["37.2"]], nrow(subset(raw_data_index_hhc_investig_arm_1, raw_data_index_hhc_investig_arm_1$tb_treat_hhm_sc_sc=='Yes')))




########################################Sample Collection
setCellValue(cells[["42.2"]], nrow(subset(raw_data_index_hhc_investig_arm_1, raw_data_index_hhc_investig_arm_1$con_out_hhm_sc=='Yes')))

setCellValue(cells[["43.2"]], nrow(subset(raw_data_index_hhc_investig_arm_1, raw_data_index_hhc_investig_arm_1$hhm_pt_sc=='Yes')))

setCellValue(cells[["44.2"]], nrow(subset(raw_data_index_hhc_investig_arm_1, raw_data_index_hhc_investig_arm_1$hhm_bio_sc=='Yes')))

setCellValue(cells[["45.2"]], nrow(subset(raw_data_index_hhc_investig_arm_1, raw_data_index_hhc_investig_arm_1$hhm_up_sc=='Yes')))

setCellValue(cells[["46.2"]], nrow(subset(raw_data_index_hhc_investig_arm_1, raw_data_index_hhc_investig_arm_1$hhm_cp_sc=='Yes')))

setCellValue(cells[["48.2"]], nrow(subset(raw_data_index_hhc_investig_arm_1, raw_data_index_hhc_investig_arm_1$hhm_sput_col_sc=='Yes' |
                                     raw_data_index_hhc_investig_arm_1$sputum_neb_sc=='Yes')))

setCellValue(cells[["49.2"]], nrow(subset(raw_data_index_hhc_investig_arm_1, raw_data_index_hhc_investig_arm_1$hhm_sput_col_sc=='No' |
                                     raw_data_index_hhc_investig_arm_1$sputum_neb_sc=='No')))


##########################################################Results
setCellValue(cells[["54.2"]], nrow(subset(raw_data_index_hhc_investig_arm_1, raw_data_index_hhc_investig_arm_1$hhci_res_immed=='Negative')))

setCellValue(cells[["55.2"]], nrow(subset(raw_data_index_hhc_investig_arm_1, raw_data_index_hhc_investig_arm_1$hhci_res_immed=='Positive')))

setCellValue(cells[["56.2"]], nrow(subset(raw_data_index_hhc_investig_arm_1, raw_data_index_hhc_investig_arm_1$hhci_res_immed=='Invalid')))

setCellValue(cells[["57.2"]], nrow(subset(raw_data_index_hhc_investig_arm_1, raw_data_index_hhc_investig_arm_1$hhci_res_immed=='No result (please specify reason in comment box)')))


#Single Swab
setCellValue(cells[["60.2"]], nrow(subset(raw_data_test_operations_arm_1, raw_data_test_operations_arm_1$to_number_swabs=='Single Swab' &
                                            (raw_data_test_operations_arm_1$to_pool_result=='Negative' |
                                               raw_data_test_operations_arm_1$to_pool_2_result=='Negative' |
                                               raw_data_test_operations_arm_1$to_pool_3_result=='Negative' |
                                               raw_data_test_operations_arm_1$to_pool_4_result=='Negative' |
                                               raw_data_test_operations_arm_1$to_pool_result_rep=='Negative' |
                                               raw_data_test_operations_arm_1$to_pool_result_rep_2=='Negative'))))



setCellValue(cells[["61.2"]], nrow(subset(raw_data_test_operations_arm_1, raw_data_test_operations_arm_1$to_number_swabs=='Single Swab' &
                                            (raw_data_test_operations_arm_1$to_pool_result=='Positive' |
                                               raw_data_test_operations_arm_1$to_pool_2_result=='Positive' |
                                               raw_data_test_operations_arm_1$to_pool_3_result=='Positive' |
                                               raw_data_test_operations_arm_1$to_pool_4_result=='Positive' |
                                               raw_data_test_operations_arm_1$to_pool_result_rep=='Positive' |
                                               raw_data_test_operations_arm_1$to_pool_result_rep_2=='Positive'))))


setCellValue(cells[["62.2"]], nrow(subset(raw_data_test_operations_arm_1, raw_data_test_operations_arm_1$to_number_swabs=='Single Swab' &
                                            (raw_data_test_operations_arm_1$to_pool_result=='Error' |
                                               raw_data_test_operations_arm_1$to_pool_2_result=='Error' |
                                               raw_data_test_operations_arm_1$to_pool_3_result=='Error' |
                                               raw_data_test_operations_arm_1$to_pool_4_result=='Error' |
                                               raw_data_test_operations_arm_1$to_pool_result_rep=='Error' |
                                               raw_data_test_operations_arm_1$to_pool_result_rep_2=='Error'))))


setCellValue(cells[["63.2"]], nrow(subset(raw_data_test_operations_arm_1, raw_data_test_operations_arm_1$to_number_swabs=='Single Swab' &
                                            (raw_data_test_operations_arm_1$to_pool_result=='Invalid' |
                                               raw_data_test_operations_arm_1$to_pool_2_result=='Invalid' |
                                               raw_data_test_operations_arm_1$to_pool_3_result=='Invalid' |
                                               raw_data_test_operations_arm_1$to_pool_4_result=='Invalid' |
                                               raw_data_test_operations_arm_1$to_pool_result_rep=='Invalid' |
                                               raw_data_test_operations_arm_1$to_pool_result_rep_2=='Invalid'))))

setCellValue(cells[["64.2"]], nrow(subset(raw_data_test_operations_arm_1, raw_data_test_operations_arm_1$to_number_swabs=='Single Swab' &
                                            (raw_data_test_operations_arm_1$to_pool_result=='No result' |
                                               raw_data_test_operations_arm_1$to_pool_2_result=='No result' |
                                               raw_data_test_operations_arm_1$to_pool_3_result=='No result' |
                                               raw_data_test_operations_arm_1$to_pool_4_result=='No result' |
                                               raw_data_test_operations_arm_1$to_pool_result_rep=='No result' |
                                               raw_data_test_operations_arm_1$to_pool_result_rep_2=='No result'))))


#Two Pooled Swab Test
setCellValue(cells[["67.2"]], nrow(subset(raw_data_test_operations_arm_1, raw_data_test_operations_arm_1$to_number_swabs=='Two Pooled' &
                                            (raw_data_test_operations_arm_1$to_pool_result=='Negative' |
                                               raw_data_test_operations_arm_1$to_pool_2_result=='Negative' |
                                               raw_data_test_operations_arm_1$to_pool_3_result=='Negative' |
                                               raw_data_test_operations_arm_1$to_pool_4_result=='Negative' |
                                               raw_data_test_operations_arm_1$to_pool_result_rep=='Negative' |
                                               raw_data_test_operations_arm_1$to_pool_result_rep_2=='Negative'))))

setCellValue(cells[["68.2"]], nrow(subset(raw_data_test_operations_arm_1, raw_data_test_operations_arm_1$to_number_swabs=='Two Pooled' &
                                            (raw_data_test_operations_arm_1$to_pool_result=='Positive' |
                                               raw_data_test_operations_arm_1$to_pool_2_result=='Positive' |
                                               raw_data_test_operations_arm_1$to_pool_3_result=='Positive' |
                                               raw_data_test_operations_arm_1$to_pool_4_result=='Positive' |
                                               raw_data_test_operations_arm_1$to_pool_result_rep=='Positive' |
                                               raw_data_test_operations_arm_1$to_pool_result_rep_2=='Positive'))))


setCellValue(cells[["69.2"]], nrow(subset(raw_data_test_operations_arm_1, raw_data_test_operations_arm_1$to_number_swabs=='Two Pooled' &
                                            (raw_data_test_operations_arm_1$to_pool_result=='Error' |
                                               raw_data_test_operations_arm_1$to_pool_2_result=='Error' |
                                               raw_data_test_operations_arm_1$to_pool_3_result=='Error' |
                                               raw_data_test_operations_arm_1$to_pool_4_result=='Error' |
                                               raw_data_test_operations_arm_1$to_pool_result_rep=='Error' |
                                               raw_data_test_operations_arm_1$to_pool_result_rep_2=='Error'))))

setCellValue(cells[["70.2"]], nrow(subset(raw_data_test_operations_arm_1, raw_data_test_operations_arm_1$to_number_swabs=='Two Pooled' &
                                            (raw_data_test_operations_arm_1$to_pool_result=='Invalid' |
                                               raw_data_test_operations_arm_1$to_pool_2_result=='Invalid' |
                                               raw_data_test_operations_arm_1$to_pool_3_result=='Invalid' |
                                               raw_data_test_operations_arm_1$to_pool_4_result=='Invalid' |
                                               raw_data_test_operations_arm_1$to_pool_result_rep=='Invalid' |
                                               raw_data_test_operations_arm_1$to_pool_result_rep_2=='Invalid'))))

setCellValue(cells[["71.2"]], nrow(subset(raw_data_test_operations_arm_1, raw_data_test_operations_arm_1$to_number_swabs=='Two Pooled' &
                                            (raw_data_test_operations_arm_1$to_pool_result=='No result' |
                                               raw_data_test_operations_arm_1$to_pool_2_result=='No result' |
                                               raw_data_test_operations_arm_1$to_pool_3_result=='No result' |
                                               raw_data_test_operations_arm_1$to_pool_4_result=='No result' |
                                               raw_data_test_operations_arm_1$to_pool_result_rep=='No result' |
                                               raw_data_test_operations_arm_1$to_pool_result_rep_2=='No result'))))


#Three Pooled Swab Test
setCellValue(cells[["75.2"]], nrow(subset(raw_data_test_operations_arm_1, raw_data_test_operations_arm_1$to_number_swabs=='Three Pooled' &
                                            (raw_data_test_operations_arm_1$to_pool_result=='Negative' |
                                               raw_data_test_operations_arm_1$to_pool_2_result=='Negative' |
                                               raw_data_test_operations_arm_1$to_pool_3_result=='Negative' |
                                               raw_data_test_operations_arm_1$to_pool_4_result=='Negative' |
                                               raw_data_test_operations_arm_1$to_pool_result_rep=='Negative' |
                                               raw_data_test_operations_arm_1$to_pool_result_rep_2=='Negative'))))

setCellValue(cells[["76.2"]], nrow(subset(raw_data_test_operations_arm_1, raw_data_test_operations_arm_1$to_number_swabs=='Three Pooled' &
                                            (raw_data_test_operations_arm_1$to_pool_result=='Positive' |
                                               raw_data_test_operations_arm_1$to_pool_2_result=='Positive' |
                                               raw_data_test_operations_arm_1$to_pool_3_result=='Positive' |
                                               raw_data_test_operations_arm_1$to_pool_4_result=='Positive' |
                                               raw_data_test_operations_arm_1$to_pool_result_rep=='Positive' |
                                               raw_data_test_operations_arm_1$to_pool_result_rep_2=='Positive'))))


setCellValue(cells[["77.2"]], nrow(subset(raw_data_test_operations_arm_1, raw_data_test_operations_arm_1$to_number_swabs=='Three Pooled' &
                                            (raw_data_test_operations_arm_1$to_pool_result=='Error' |
                                               raw_data_test_operations_arm_1$to_pool_2_result=='Error' |
                                               raw_data_test_operations_arm_1$to_pool_3_result=='Error' |
                                               raw_data_test_operations_arm_1$to_pool_4_result=='Error' |
                                               raw_data_test_operations_arm_1$to_pool_result_rep=='Error' |
                                               raw_data_test_operations_arm_1$to_pool_result_rep_2=='Error'))))

setCellValue(cells[["78.2"]], nrow(subset(raw_data_test_operations_arm_1, raw_data_test_operations_arm_1$to_number_swabs=='Three Pooled' &
                                            (raw_data_test_operations_arm_1$to_pool_result=='Invalid' |
                                               raw_data_test_operations_arm_1$to_pool_2_result=='Invalid' |
                                               raw_data_test_operations_arm_1$to_pool_3_result=='Invalid' |
                                               raw_data_test_operations_arm_1$to_pool_4_result=='Invalid' |
                                               raw_data_test_operations_arm_1$to_pool_result_rep=='Invalid' |
                                               raw_data_test_operations_arm_1$to_pool_result_rep_2=='Invalid'))))

setCellValue(cells[["79.2"]], nrow(subset(raw_data_test_operations_arm_1, raw_data_test_operations_arm_1$to_number_swabs=='Three Pooled' &
                                            (raw_data_test_operations_arm_1$to_pool_result=='No result' |
                                               raw_data_test_operations_arm_1$to_pool_2_result=='No result' |
                                               raw_data_test_operations_arm_1$to_pool_3_result=='No result' |
                                               raw_data_test_operations_arm_1$to_pool_4_result=='No result' |
                                               raw_data_test_operations_arm_1$to_pool_result_rep=='No result' |
                                               raw_data_test_operations_arm_1$to_pool_result_rep_2=='No result'))))



####################################Lab Results
setCellValue(cells[["83.2"]], nrow(subset(raw_data_index_hhc_investig_arm_1, raw_data_index_hhc_investig_arm_1$up_res_shipped=='Yes')))

setCellValue(cells[["85.2"]], nrow(subset(raw_data_index_hhc_investig_arm_1, raw_data_index_hhc_investig_arm_1$up_res=='Negative')))

setCellValue(cells[["86.2"]], nrow(subset(raw_data_index_hhc_investig_arm_1, raw_data_index_hhc_investig_arm_1$up_res=='Positive')))



setCellValue(cells[["91.2"]], mean(raw_data_household_level_da_arm_1$invest_time, na.rm = TRUE))

setCellValue(cells[["92.2"]], median(raw_data_household_level_da_arm_1$invest_time, na.rm = TRUE))

setCellValue(cells[["93.2"]], min(raw_data_household_level_da_arm_1$invest_time, na.rm = TRUE))

setCellValue(cells[["94.2"]], max(raw_data_household_level_da_arm_1$invest_time, na.rm = TRUE))



setCellValue(cells[["98.2"]], mean(raw_data_test_operations_arm_1$to_temp, na.rm = TRUE))

setCellValue(cells[["99.2"]], median(raw_data_test_operations_arm_1$to_temp, na.rm = TRUE))

setCellValue(cells[["100.2"]], min(raw_data_test_operations_arm_1$to_temp, na.rm = TRUE))

setCellValue(cells[["101.2"]], max(raw_data_test_operations_arm_1$to_temp, na.rm = TRUE))


xlsx::saveWorkbook(wb1, filename_new)

