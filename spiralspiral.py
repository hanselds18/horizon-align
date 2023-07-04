import math
from mpmath import sec

class SpiralSpiral():
    def ss(distance_hor, delta_azimuth_hor, speed_cell, elevation_normal_hor, elevation_max_hor, radius_cell, one_lane_width, rotated_lane,\
           total_lane, dict_1, dict_2, dict_3, dict_4, dict_5, dict_6, dict_7, dict_8, dict_9, dict_10, dict_11, dict_12, dict_13,dict_14, turn_direction):

        # ----------------------------------------------------------- Spiral - Spiral ---------------------------------------------------#
        if int(radius_cell) < dict_1[int(speed_cell)][int(elevation_max_hor)] and int(radius_cell) < dict_2[int(speed_cell)][0] and int(radius_cell) >= dict_4[int(speed_cell)][
            int(elevation_max_hor)]:
            # print('Pakai SS')
            curve_type = 'SS'
            # one_lane_width  = input('Masukkan lebar 1 lajur (meter)           : ')
            # rotated_lane    = input('Masukkan jumlah lajur yang di rotasi (buah) : ')
            bw              = dict_5[int(rotated_lane)][(0)]
            delta           = dict_6[int(speed_cell)][(0)]

            if elevation_normal_hor == -2 and elevation_max_hor == 8:
                elevation_design = round(dict_7[int(radius_cell)][int(speed_cell)], 1)
                length_spiral_kelandaian = round((float(one_lane_width) * float(rotated_lane) * elevation_design / 100 * bw) / delta, 1)
            elif elevation_normal_hor == -2 and elevation_max_hor == 6:
                elevation_design = dict_8[int(radius_cell)][int(speed_cell)]
                length_spiral_kelandaian = round((float(one_lane_width) * float(rotated_lane) * elevation_design / 100 * bw) / delta, 1)
            elif elevation_normal_hor == -2 and elevation_max_hor == 4:
                elevation_design = dict_9[int(radius_cell)][int(speed_cell)]
                length_spiral_kelandaian = round((float(one_lane_width) * float(rotated_lane) * elevation_design / 100 * bw) / delta, 1)

            tetha_s                     = round(float(delta_azimuth_hor) / 2, 1)
            length_spiral_offset        = round(math.sqrt(24 * 0.2 * int(radius_cell)), 1)
            length_spiral_acceleration  = round((0.0214 * (float(speed_cell) ** 3)) / (float(radius_cell) * 1.2), 1)
            length_spiral_good          = round((math.pi * int(radius_cell) * tetha_s) / 90, 1)

            # print('\nLS Kelandaian Relatif\t = ' + str(length_spiral_kelandaian) + ' meter')
            # print('LS Offset Lateral\t\t = ' + str(length_spiral_offset) + ' meter')
            # print('LS Akselerasi Lateral\t = ' + str(length_spiral_acceleration) + ' meter')
            # print('LS SS Bagus\t\t\t = ' + str(length_spiral_good) + ' meter')

            #-----------------------Cek apakah LS SS Bagus adalah yang paling besar ---------------------------------#
            if length_spiral_good == max(length_spiral_kelandaian, length_spiral_offset, length_spiral_acceleration, length_spiral_good):
                tangent_runout = round(abs(float(elevation_normal_hor) / 100) / (elevation_design / 100) * max(length_spiral_kelandaian,\
                                                                                                         length_spiral_offset, length_spiral_acceleration, length_spiral_good), 1)
                # print('Tangent Runout\t\t\t = ' + str(tangent_runout) + ' meter')

                if speed_cell < 80:
                    lrr = round((elevation_design / 100 - float(elevation_normal_hor) / 100) / (3.6 * 0.035) * float(speed_cell), 1)
                else:
                    lrr = round((elevation_design / 100 - float(elevation_normal_hor) / 100) / (3.6 * 0.025) * float(speed_cell), 1)

                # print('\nLrr = ' + str(lrr))


                if lrr <= max(length_spiral_kelandaian, length_spiral_offset, length_spiral_acceleration, length_spiral_good) + tangent_runout:
                    # print('Lrr \u2264 Ls + TR(Lt)\t HASIL OK')
                    length_spiral_design = max(length_spiral_kelandaian, length_spiral_offset, length_spiral_acceleration, length_spiral_good)
                    if length_spiral_design < dict_10[int(speed_cell)][(0)]:
                        return "false:Ls < Ls min"
                else:
                    # print('Lrr \u003E Ls + TR(Lt)\t HASIL NOT OK')
                    return "false:Lrr > Ls + TR(Lt)"

                length_spiral_total = round(2 * length_spiral_good, 1)

                if length_spiral_total < dict_3[int(speed_cell)][0]:
                    # print('Ltotal < Lmin !! Perhitungan salah !')
                    return "false:Ltotal < Lmin"

                xs = round(length_spiral_design * (1 - (length_spiral_design ** 2 / (40 * int(radius_cell) ** 2)) + (length_spiral_design ** 4 / (3456 * int(radius_cell) ** 4))), 1)
                ys = round((length_spiral_design ** 2 / (6 * int(radius_cell))) * (1 - (length_spiral_design ** 2 / (56 * int(radius_cell) ** 2) + (length_spiral_design ** 4 / (7040 * int(radius_cell) ** 4)))), 1)
                k = round(xs - int(radius_cell) * math.sin(math.radians(tetha_s)), 1)
                p = round(ys - int(radius_cell) * (1 - math.cos(math.radians(tetha_s))), 1)
                ts = round(((int(radius_cell) + p) * math.tan(math.radians(float(delta_azimuth_hor) / 2)) + k), 1)
                es = round(((int(radius_cell) + p) * sec(math.radians(float(delta_azimuth_hor) / 2)) - int(radius_cell)), 1)

                ts_tr = ts + tangent_runout

                #-------------------Widening------------------#
                widening = dict_14[int(radius_cell)][int(speed_cell)]
                # if vehicle_type == 'TRUK TUNGGAL':
                #     widening = dict_13[int(radius_cell)][0] * total_lane
                #
                # elif vehicle_type == 'TRUK SEMI TRAILER':
                #     widening = dict_13[int(radius_cell)][1] * total_lane
                #
                # else:
                #     widening = None


                return [curve_type, float(radius_cell), float(dict_2[int(speed_cell)][0]), float(elevation_design), length_spiral_kelandaian,
                        length_spiral_offset, length_spiral_acceleration, lrr, length_spiral_good, float(dict_10[int(speed_cell)][(0)]), length_spiral_design, tetha_s, None, None,
                        length_spiral_total, float(dict_3[int(speed_cell)][0]), xs, ys, k, p, None, None, ts, tangent_runout, ts_tr, '', \
                        '', None, None, None, None, '', '', float(widening), '', '', [], [], [], [], [], [], [], [], '', turn_direction]
            else:
                return "false:The middle of the 2 spirals will snap (not connected) because Ls SS is not the greatest value."

        else:
            return "false:Radius cannot be designed as a Spiral-Spiral because its radius is not fulfill the conditions. Please change either the radius min & max or the curve type option on Ms. Excel"