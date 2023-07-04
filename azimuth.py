import math
from simple_colors import *

class Azimuth():
    def coor(nCoordinate, arrX, arrY):
        data = {}
        x_before = arrX[0]
        y_before = arrY[0]

        azimuth_before = 0

        for i in range(1, nCoordinate):
            data[i-1] = {}

            x_after   = arrX[i]
            y_after   = arrY[i]

            if i > 0:
                azimuth_after   = round(math.degrees(math.atan((x_after - x_before) / (y_after - y_before))), 1)

                if (x_after - x_before) >= 0 and (y_after - y_before) >= 0:
                    quadrant        = 1
                elif (x_after - x_before) >= 0 and (y_after - y_before) < 0:
                    quadrant        = 2
                    azimuth_after   += 180
                elif (x_after - x_before) < 0 and (y_after - y_before) < 0:
                    quadrant        = 3
                    azimuth_after   += 180
                elif (x_after - x_before) < 0 and (y_after - y_before) >= 0:
                    quadrant        = 4
                    azimuth_after   += 360

                delta       = abs(azimuth_before - azimuth_after)
                distance    = round(math.sqrt((abs(x_after - x_before)**2) + abs(y_after - y_before)**2), 1)

                if azimuth_after < azimuth_before:
                    turn = 'LEFT'
                elif azimuth_after > azimuth_before:
                    turn = 'RIGHT'
                else:
                    turn = 'STRAIGHT'

    #---------------------------------------------- Print Section ----------------------------------------
                # print('Kuadran\t\t = ' + str(quadrant))
                data[i-1]['kuadran'] = str(quadrant)

                # print('Azimuth\t\t = ' + str(azimuth_after) + ' degree')
                data[i-1]['azimuth'] = str(azimuth_after)

                if(i > 1):
                    # print('\u0394\t\t\t = ' + str(delta) + ' degree')
                    data[i-1]['delta'] = str(delta)

                # print('Distance\t = ' + str(distance) + ' meter')
                data[i-1]['distance'] = str(distance)

                if(i > 1):
                    # print('PI ' + str(i - 1) + ' turn ' + turn + '\n')
                    data[i-1]['pi'] = 'Turn\t: ' + turn
    #------------------------------------------------------------------------------------------------------
                x_before        = x_after
                y_before        = y_after
                azimuth_before = azimuth_after

        return data