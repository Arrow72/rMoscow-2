setwd('C:/.....') # ВСТАВИТЬ СВОЙ КАТАЛОГ !!!
library(XLConnect)

# Формат чисел
pp<- function (x, ...) 
{ 
  a<-ifelse(is.na(x),'',format(round(x), ..., big.mark = " ", scientific = FALSE, trim = TRUE))
  return(a)
}

pp_proc<- function (x, ...) 
{ 
  a<-ifelse(is.na(x),'',format(round(x*100, digits=2), ..., big.mark = " ", scientific = FALSE, trim = TRUE))
  return(paste0(a,'%'))
}


# Загрузили данные из файла
wb <- loadWorkbook('example.xlsx')
daf <- readWorksheet(wb, startRow=1,startCol=1, endRow=8, endCol=350, header=FALSE, sheet = 1) 

# Выбрали периоды
dates <-factor(as.Date(na.omit(t(daf[1,]))))

# Выбрали рынки
markets <- daf[4:8,1]

# Заготовка для нормальных данных
normal_df <- data.frame(period=NA, market=NA, 
                        sale_pln=NA, sale_fct=NA, 
                        revenue_pln=NA, revenue_fct=NA, 
                        margin_pln=NA, margin_fct=NA)
normal_df <- normal_df[-1,]

# Цикл по датам (порядковый номер)
for(i in 1:length(dates))
{
  for(j in 1:length(markets)) 
  {
    normal_df <- rbind(normal_df,
                  data.frame(period=dates[i], market=markets[j],
                             sale_pln=daf[3+j, 2+(i-1)*6],
                             sale_fct=daf[3+j, 3+(i-1)*6],
                             revenue_pln=daf[3+j, 4+(i-1)*6],
                             revenue_fct=daf[3+j, 5+(i-1)*6],
                             margin_pln=daf[3+j, 6+(i-1)*6],
                             margin_fct=daf[3+j, 7+(i-1)*6]
                            )
                  )
  }
  
}

# Закачивали все строками, поэтому переведем в числа
normal_df[,3] <- as.numeric(gsub(",", "", normal_df[,3], fixed = TRUE))
normal_df[,4] <- as.numeric(gsub(",", "", normal_df[,4], fixed = TRUE))
normal_df[,5] <- as.numeric(gsub(",", "", normal_df[,5], fixed = TRUE))
normal_df[,6] <- as.numeric(gsub(",", "", normal_df[,6], fixed = TRUE))

normal_df[,7] <- as.numeric(sub("%","",sub(",", ".", normal_df[,7], fixed = TRUE),fixed = TRUE))/100
normal_df[,8] <- as.numeric(sub("%","",sub(",", ".", normal_df[,8], fixed = TRUE),fixed = TRUE))/100

# Создадим новых переменных

library(lubridate)
normal_df$year  <- factor(year(normal_df$period))     #ГОД
normal_df$month <- factor(month(normal_df$period))    #МЕСЯЦ
normal_df$d_plan <- normal_df$sale_fct/normal_df$sale_pln #ВЫПОЛНЕНИЕ ПЛАНА

# Создадим еще табличку aka - PivotTable (сводная таблица)
library(dplyr)
library(reshape2)
library(ReporteRs)

ft <- as.data.frame(select(normal_df, year, month, sale_fct) 
                    %>% group_by(year, month) %>% summarise(sale_fct=sum(sale_fct)))
ft <- dcast(ft, month~year,sum)
colnames(ft) <- c('Месяц','2015','2016')
ft$Прирост=ft$`2016`/ft$`2015`-1
ft[,2:3] <- pp(ft[,2:3])
ft[,4] <- pp_proc(ft[,4])

# Сохраним индексы знаков приростов
plus_ind <- as.numeric(ft[which(ft$Прирост>0),1])
minus_ind<- as.numeric(ft[which(ft$Прирост<0),1])





library(ggplot2)
# 
# Рисуем графики - сразу в функции
Line_graph_fun <- function()
{
  print(
    ggplot(data=normal_df[normal_df$market=='ИТОГО',c('period', 'sale_fct','month','year')])+
      ggtitle('Отгрузка 2015-2016', subtitle = 'по месяцам')+
      geom_line(aes(x=month, y=sale_fct,group=year, color=year))+
      scale_y_continuous(labels=pp, name='Отгрузка, руб', limits = c(0, 3000000))+
      scale_x_discrete(name='Месяцы')+
      theme_bw()
  )
}


plan_graph_fun <- function()
{
  print(
    ggplot()+
      geom_bar(data=normal_df[
        normal_df$market=='ИТОГО'&
          normal_df$year==2015,
        ], 
        aes(x=period,y=d_plan, fill=factor(sign(d_plan-1))), 
        stat='identity', alpha=0.25)+
      geom_bar(data=normal_df[
        normal_df$market=='ИТОГО'&
          normal_df$year==2016,
        ], 
        aes(x=period,y=d_plan, fill=factor(sign(d_plan-1))), 
        stat='identity', alpha=0.80)+
      
      scale_x_discrete(labels=normal_df[normal_df$market=='ИТОГО',]$month, name='месяц')+
      scale_y_continuous(labels=NULL, name='Выполнение плана, %')+
      theme_bw()+
      scale_fill_manual(values=c("indianred2", "forestgreen"), guide=FALSE)+
      geom_text(data=normal_df[normal_df$market=='ИТОГО',], 
                aes(label=pp_proc(d_plan),x=period, y=d_plan), angle=90, fontface="bold", hjust=1.2)
  )
}

lfl_graph_fun <- function()
{
  print(
    ggplot()+
      geom_bar(data=normal_df[normal_df$market!='ИТОГО',], aes(x=year,y=sale_fct, fill=market),  stat="identity")+
      facet_wrap(~month,nrow=1, strip.position = 'bottom')+
      theme(
        panel.border=element_blank(),
        panel.background = element_blank(),
        panel.grid.major=element_blank(),
        panel.grid.minor = element_blank(), 
        # panel.margin = unit(0.5, "lines"),
        panel.spacing = unit(0.5, "lines"),
        plot.background = element_blank()
      )+
      theme(axis.text.x = element_text(angle=90, hjust=1, vjust=.5, size=7.5))+
      theme(axis.title.x=element_blank())+
      scale_y_continuous(labels=pp, name='Отгрузка')
  )
}

# И таблетку
library(dplyr)
library(magrittr)


pie_graph <- function(inYear)
{
  t_y <- inYear
  d <- as.data.frame(select(normal_df[as.numeric(as.character(normal_df$month))<11 & normal_df$year==t_y & normal_df$market!='ИТОГО',],
                            year, market, sale_fct) %>% group_by(year,market) %>% summarise(sum(sale_fct)))
  colnames(d) <- c('year', 'market', 'sale_fct')
  d$sale_fct <- d$sale_fct*-1
  d <- d %>%
    mutate(market = factor(market, levels = market)) %>%
    arrange(desc(market)) %>%
    mutate(cum_y = cumsum(sale_fct),
           centres = cum_y - sale_fct / 2,
           lab = sale_fct/sum(d$sale_fct)
    )

  d$market <- as.character(d$market)
  d$cum_y <- cumsum(d$sale_fct) - d$sale_fct/2
  d$lab <- d$sale_fct/sum(d$sale_fct)
  leg <- guide_legend(reverse = FALSE)
  plot(
    ggplot(d, aes(x = 1, weight = sale_fct, fill = market)) +
      ggtitle(t_y)+
      geom_bar(width = 1)+
      coord_polar(theta = "y") +
      geom_text(x = 1.1, aes(y = centres, label = pp_proc(lab)),size=7, color='white', fontface = "bold") +
      theme(
        panel.grid=element_blank(),
        legend.title =element_blank(),
        legend.position ='bottom',
        axis.text=element_blank(),
        axis.ticks=element_blank(),
        axis.title = element_blank(),
        panel.grid.minor = element_blank(), 
        panel.grid.major = element_blank(),
        panel.background = element_blank(),
        panel.border = element_blank(),
        plot.background = element_blank()
      ) 
  )
}

pie_graph_2015 <- function()
{
  pie_graph(2015)
}
pie_graph_2016 <- function()
{
  pie_graph(2016)
}


ft <- FlexTable(data = ft, add.rownames = FALSE, header.columns = TRUE,
          body.text.props = textProperties( font.size = 12 ),
          header.text.props = textProperties( font.size = 13, font.weight = "bold"))

setFlexTableBackgroundColors(ft, minus_ind,4, 'sienna2', to="body")
setFlexTableBackgroundColors(ft, plus_ind,4, 'lightgreen', to="body")

# И создаем презентацию


presentation_title <- c('Сравнительные итоги 2015-2016г.')
presentator_name <- 'Автор: Выбегалло А.А.'
options('ReporteRs-fontsize'= 18, 'ReporteRs-default-font'='Tahoma')
doc <- pptx(template="SaleReport_Template.pptx" ) 

# Slide 1 : Title slide
#+++++++++++++++++++++++
doc <- addSlide(doc, "Титульный слайд")
doc <- addTitle(doc,presentation_title)
doc <- addSubtitle(doc, presentator_name)

# Slide 2 : Линейный график и таблицы
#+++++++++++++++++++++++
doc <- addSlide(doc, "Только заголовок")
doc <- addTitle(doc,'Динамика отгрузок')

doc <- addPlot(doc, Line_graph_fun ,  offx = 0, offy = 2, width = 6, height = 5)
doc <- addFlexTable(doc, ft, offx = 6, offy = 3, width = 3.5, height = 3,
                    par.properties = #parProperties(text.align = "center"))
                      parProperties( text.align = "center", 
                                     padding.top = 5,
                                     padding.bottom = 5,
                                     padding.left = 2,
                                     padding.right = 2
                      )
)



# Slide 3 : Разноцветной график
#+++++++++++++++++++++++
doc <- addSlide(doc, "Только заголовок")
doc <- addTitle(doc,'Анализ выполнения плана')
doc <- addPlot(doc, plan_graph_fun,  offx = 0, offy = 2, width = 10, height = 5.5)

# Slide 4 : lfl график
#+++++++++++++++++++++++
doc <- addSlide(doc, "Только заголовок")
doc <- addTitle(doc,'Like-for-like анализ рынков')
doc <- addPlot(doc, lfl_graph_fun,  offx = 0, offy = 2, width = 10, height = 5.5)


# Slide 5 : pie -диаграммы
#+++++++++++++++++++++++
doc <- addSlide(doc, "Только заголовок")
doc <- addTitle(doc,'Доля рынков в отгрузке')
doc <- addPlot(doc, pie_graph_2015,  offx = 0, offy = 2, width = 5, height = 5)
doc <- addPlot(doc, pie_graph_2016,  offx = 5, offy = 2, width = 5, height = 5)


# Сохраняем нашу презентацию
writeDoc(doc, "Ready_present.pptx" )
