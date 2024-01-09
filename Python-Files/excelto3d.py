import math
import open3d as o3d
import numpy as np
import xlrd

book = xlrd.open_workbook("dataset.xls")
print("The number of worksheets is {0}".format(book.nsheets))
print("Worksheet name(s): {0}".format(book.sheet_names()))
sh = book.sheet_by_index(0)
print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))

xyz = open("points.xyz", "w")
xyz.close
for dcol in range (2,sh.ncols):
    x = float(sh.cell_value(0, dcol)) / 1000
    for i in range(1,513):
        distance = float(sh.cell_value(i, dcol))
        angle = math.radians(float(sh.cell_value(i, 1)))
        y = distance * (math.cos(angle)) / 1000
        z = distance * (math.sin(angle)) / 1000
        xyz = open("points.xyz", "a")
        xyz.write("{} {} {}\n".format(x, y, z))
        xyz.close
mesh = []
for i in range(512):
    for j in range(sh.ncols-3):
        if i == 511 and j == sh.ncols-4:
            break
        mesh.append([j*512+i,(j+1)*512+i])

total_points = (sh.ncols-2) * 512
for i in range(total_points-2):
    mesh.append([i,i+1])
    
pcd = o3d.io.read_point_cloud("points.xyz", format='xyz')

line_set = o3d.geometry.LineSet(points=o3d.utility.Vector3dVector(np.asarray(pcd.points)),
                                lines=o3d.utility.Vector2iVector(mesh))
o3d.visualization.draw_geometries([line_set])

