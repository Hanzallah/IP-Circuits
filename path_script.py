# -*- coding: utf-8 -*-
"""
@authors: Hanzallah Azim Burney, Abdul Hamid Daboussi and Shabnam Sadigova
"""
import sys
import xlrd
import xlwt

def get_matrix(points, rectangles):
    cost = []
    for i in range(len(points)):
        cost.append([])
        for j in range(len(points)):
            cost[i].append(0)

    for i in range(len(points)):
        for j in range(len(points)):
            if i == j:
                cost[i][j] = sys.maxsize
            else:
                cost[i][j] = abs(points[i].x - points[j].x) + abs(points[i].y - points[j].y)
    
    for i in range(len(points)):    
        for j in range(len(points)):
            flag1 = False
            flag2 = False
            for rectangle in rectangles:                        
                if  (((points[i].y >= rectangle.y + rectangle.height) and (points[j].y <= rectangle.y))\
                    and ((points[i].x >= rectangle.x) and (points[i].x <= (rectangle.x + rectangle.width))))\
                    or  ((points[i].x <= rectangle.x) and (points[j].x >= rectangle.x + rectangle.width)\
                    and ((points[j].y >= rectangle.y) and (points[j].y <= (rectangle.y + rectangle.height)))):
                    if flag2:
                       cost[i][j] = sys.maxsize 
                       cost[j][i] = sys.maxsize 
                       break
                    flag1 = True
                    
                if  (((points[i].x <= rectangle.x) and (points[j].x >= rectangle.x + rectangle.width))\
                    and ((points[i].y >= rectangle.y) and (points[i].y <= (rectangle.y + rectangle.height))))\
                    or  ((points[i].y >= (rectangle.y + rectangle.height)) and (points[j].y <= rectangle.y)\
                    and ((points[i].x >= rectangle.x) and (points[i].x <= (rectangle.x + rectangle.width)))):
                    if flag1:
                       cost[i][j] = sys.maxsize 
                       cost[j][i] = sys.maxsize 
                       break
                    flag2 = True
    return cost


class rectangle:
    def __init__(self, x = 0, y = 0, width = 0, height = 0):
        self.x = x
        self.y = y
        self.width = width
        self.height = height
    
class point:
    def __init__(self, x = 0, y = 0):
        self.x = x
        self.y = y

    def __ne__(self, point):
        if self.x != point.x and self.y != point.y:
            return True
        return False


def get_excel(pathname):
    workbook = xlrd.open_workbook(pathname)
    points_data = workbook.sheet_by_index(0)
    rectangles_data = workbook.sheet_by_index(1)
    
    points = []
    rectangles = []
    
    try:
        i = 1
        j = 0
        while points_data.cell(i,j) != xlrd.empty_cell.value:
            x = int(points_data.cell(i,j).value)
            y = int(points_data.cell(i,j+1).value)
            points.append(point(x,y))
            i += 1
    except:
        print("Finished")
    
    try:
        i = 2
        j = 0
        
        while rectangles_data.cell(i,j) != xlrd.empty_cell.value:
            x = int(rectangles_data.cell(i,j).value)
            y = int(rectangles_data.cell(i,j+1).value)
            width = int(rectangles_data.cell(i,j+2).value) 
            height = int(rectangles_data.cell(i,j+3).value)
            rectangles.append(rectangle(x,y, width, height))
            i += 1
    except:
        print("Finished")
      
    return (points, rectangles)
        

def write_to_excel(cost):
    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("Python Sheet 1") 
    for i in range(len(cost)):
        row = sheet1.row(i)
        for j in range(len(cost[i])):
            row.write(j, cost[i-1][j-1])
    book.save("temp.xls")
    

points, rectangles = get_excel(r"data2-nazlı.xlsx")

cost = get_matrix(points,rectangles)
write_to_excel(cost)