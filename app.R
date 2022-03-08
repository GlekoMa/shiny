# 命令行输入：
# options(encoding = "UTF-8")
# rsconnect::deployApp('C:/Users/Gleko/Desktop/R_rela/score_app')

# Packages-------------------------------

library(shiny)
library(readxl)
library(stringr)
library(openxlsx)
library(reshape2)
library(dplyr)

# Shiny----------------------------------

ui <- fluidPage(
    fileInput("upload", "Require 'xxx.xlsx'"),
    downloadButton("download", "Download"),
    tableOutput("files")
)

server <- function(input,output,session){
    output$files <- renderTable({
        req(input$upload)
        
        # Prepare-------------------------------------------------------
        
        # 读取数据
        score0 <- read_excel(input$upload$datapath)
        # 删除无价值的行和列（第一行、最后两行;第三列、最后七列）
        score0 <- score0[c(-1, -nrow(score0), 1-nrow(score0)), c(-3, (-ncol(score0)+6):-ncol(score0))]
        # 对score0作备份以供生成Sheet2使用
        sc0_back_up <- as.data.frame(score0)
        # 把课程信息格式改为“课程名 学分”；将数据框第一行作为列名
        ## 按“，”把课程信息中的词分离
        course <- as.data.frame(do.call(rbind, str_split(string = score0[1, -1:-2], pattern = "，")))
        course[, 4] <- paste('/', do.call(rbind, str_split(string = course[, 4], pattern = ' '))[, 2], sep = '')
        ## 取出格式化全部课程
        all_cour <- course[, c(1, 4)]
        all_cour <- str_c(course[, 1], course[, 4], sep = "")
        ## 取出格式化必修课程
        ### 去掉由各种原因导致成绩缺失的非目标列
        ### 如：转专业同学的原专业必修课
        col_na <- rep(NA, ncol(score0)-2)
        for (i in 3:(length(col_na)+2)){
            col_na[i-2] <- score0[,i] %>% 
                is.na() %>% 
                sum()
        } # 对每列缺失值NA计数
        require_cour <- course[col_na / nrow(score0) < 0.5, ]
        require_cour <- require_cour[require_cour[,3] == "必修", c(1, 4)]
        require_cour <- str_c(require_cour[, 1], require_cour[, 2], sep = "")
        ## 将第一行作为列名
        names(score0) <- c("学号", "姓名", all_cour)
        score0 <- as.data.frame(score0[-1, ])
        ## 将character类型数字转换为numeric类型
        score0[, -2] <- apply(score0[, -2], 2, as.numeric)
        
        # Sheet1--------------------------------------------------------
        
        sheet1 <- score0[, c(1, 2, which(names(score0) %in% require_cour))]
        # 学习成绩＝（科目一*对应学分+科目二*对应学分···+科目n*对应学分）/总学分
        ## 分离出每门课的学分，生成(ncol-2)×1矩阵mat_point
        cour_point <- do.call(rbind, str_split(string = require_cour, pattern = "/"))
        mat_point <- apply(matrix(cour_point[, 2]), 2, as.numeric)
        ## 由课程分数生成nrow×(ncol-2)矩阵mat_score
        mat_score <- apply(as.matrix(sheet1[, -1:-2]), 2, as.numeric)
        ## 算出学习成绩
        sheet1$学习成绩 <- format(as.numeric(mat_score %*% mat_point / sum(mat_point[, 1])), digits = 4)
        ## 按成绩排序
        sheet1 <- arrange(sheet1, desc(学习成绩))
        ranking <- data.frame("排名" = 1:nrow(sheet1))
        ## 生成成绩表Sheet1
        sheet1 <- cbind(sheet1[, c(1, 2)], ranking, sheet1[, 3:ncol(sheet1)])
        
        # Sheet2--------------------------------------------------------
        
        ## 因某些原因（如转专业），同名课程可能被分成两列，需将重复的列重命名为“原列名 ”
        dup_course <- duplicated(as.character(sc0_back_up[1, ]))
        if (sum(dup_course) != 0){
            sc0_back_up[1, dup_course] <- str_c(sc0_back_up[1, dup_course], ' ')
        }
        names(sc0_back_up) <- sc0_back_up[1, ]
        sc0_back_up <- sc0_back_up[-1, ]
        ## 长短列表转换
        sheet2 <- melt(sc0_back_up, id.vars = c("学号", "姓名"),
                       variable.name = "科目",
                       value.name = "分数")
        ## 删去缺失值的行
        sheet2 <- sheet2[!is.na(sheet2$分数), ]
        ## 筛选出成绩小于60的
        sheet2 <- sheet2[sheet2$分数 < 60 & sheet2$分数 != 100, ]
        ## 把课程列细分成科目类型和科目
        cour_df <- data.frame(do.call(rbind, str_split(string = sheet2[, 3], pattern = "，")))
        sheet2[,3] <- cour_df[, 1]
        sheet2$科目类型 <- cour_df[, 2]
        ## 生成挂科表sheet2
        sheet2 <- sheet2[,c(1,2,5,3,4)]
        
        # Sheet3--------------------------------------------------------
        
        ## 将sheet1长短列表转换
        fail_cour <- melt(sheet1[, c(-3, -ncol(sheet1))],
                          na.rm = TRUE,
                          id.vars = c("学号", "姓名"),
                          variable.name = "科目",
                          value.name = "分数")
        ## 计每个科目挂科数
        fail_cour <- fail_cour[fail_cour$分数 < 60, ] %>% group_by(科目) %>% count()
        ## 新添加一列“挂科率”
        fail_cour$挂科率 <- format(fail_cour$n/nrow(sheet1), digit = 1)
        ## 转置数据框以待之后列合并
        fail_cour <- as.data.frame(t(fail_cour))
        ## 第一行作为列名
        names(fail_cour) <- fail_cour[1, ]
        fail_cour <- fail_cour[-1, ]
        ## 创建一个数据框，表示班级挂科率
        n <- rep(0,length(require_cour))
        sheet3 <- as.data.frame(rbind(require_cour, n, n))
        sheet3$new <- c("科目","挂科人数","挂科率")
        sheet3 <- sheet3[, c(ncol(sheet3), 1:(ncol(sheet3)-1))]
        names(sheet3) <- sheet3[1,]
        sheet3 <- sheet3[-1,]
        ## 赋值
        sheet3 <- cbind(sheet3[, !names(sheet3) %in% names(fail_cour)], fail_cour)
        
        # Create Excel--------------------------------------------------
        
        wb <- createWorkbook(creator = 'Rydeen', title = '成绩')
        addWorksheet(wb,sheetName = '学习成绩')
        writeData(wb,sheet = '学习成绩',x = sheet1)
        addWorksheet(wb,sheetName = '挂科名单')
        writeData(wb,sheet = '挂科名单',x = sheet2)
        addWorksheet(wb,sheetName = '班级挂科率')
        writeData(wb,sheet = '班级挂科率',x = sheet3)
        saveWorkbook(wb, "score_server.xlsx", overwrite = TRUE)
        
        return(sheet1)
        rm(list = ls())
    })
    output$download <- downloadHandler(
        filename = "成绩单(beta).xlsx",
        content = function(file) {
            file.copy(paste0(getwd(),"/score_server.xlsx"), file)
        }
    )
}

shinyApp(ui,server)