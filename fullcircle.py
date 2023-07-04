import math
from mpmath import sec
from jarakbebassamping import JarakBebasSamping

class FullCircle():
    def fc(distance_hor, delta_azimuth_hor, speed_cell, elevation_normal_hor, elevation_max_hor,radius_cell, one_lane_width, rotated_lane,\
            total_lane, dict_1, dict_2, dict_3, dict_4, dict_5, dict_6, dict_7, dict_8, dict_9, dict_10, dict_11, dict_12, dict_13, dict_14, ruwasja,\
           rs, psd, turn_direction):
        # ---------------------------------------------------- Full Circle without Superelevation ------------------------------------------------------------------#
        global m
        if int(radius_cell) >= dict_1[int(speed_cell)][int(elevation_max_hor)] and int(radius_cell) > dict_2[int(speed_cell)][0]:
            # print(dict_1[int(speed_cell)][int(elevation_max_hor)])
            # print(dict_2[int(speed_cell)][0])
            # print('Pakai FC tanpa Elevasi')
            curve_type = 'FC without Elevation'
            length_curve = round(2 * math.pi * (int(radius_cell)) * (float(delta_azimuth_hor)) / 360, 1)
            if dict_3[int(speed_cell)][0] == KeyError:
                return "false:Speed < 30 km/h. Check Ms. Excel"
            elif length_curve < dict_3[int(speed_cell)][0]:
                return "false:Lc < Lmin"

                # print('LC < Lmin !! Perhitungan salah !')

            # print('LC = ' + str(length_curve))
            tangent_length = round(int(radius_cell) * math.tan(math.radians(float(delta_azimuth_hor) / 2)), 1)
            # print('T  = ' + str(tangent_length))
            external_distance = round(int(radius_cell) * (sec(math.radians(float(delta_azimuth_hor) / 2)) - 1), 1)
            # print('E  = ' + str(external_distance))
            middle_ordinate = round(int(radius_cell) * (1 - math.cos(math.radians(float(delta_azimuth_hor) / 2))), 1)
            # print('M  = ' + str(middle_ordinate) + '\n')

            # ------ Jarak Bebas Samping -----#
            if psd == "Y":
                ssd = dict_11[int(speed_cell)][(0)]
                psd = dict_12[int(speed_cell)][(0)]
                sight_distance = max(ssd, psd)

                if rs == None:
                    return "false: Rs is blank. Please check your Ms. Excel"

                elif ruwasja == None:
                    return "false: HSO is blank. Please check your Ms. Excel"

                elif sight_distance <= length_curve:
                    m = round(float(rs) * (1 - math.cos(math.radians(28.65 * sight_distance) / float(rs))), 1)

                    if float(ruwasja) < m:
                        return "false:Not Enough Horizontal Sightline Offset Space (M space < M required)"


                    else:
                        widening = dict_14[int(radius_cell)][int(speed_cell)]
                        # if vehicle_type == 'TRUK TUNGGAL':
                        #     widening = dict_13[int(radius_cell)][0] * total_lane
                        #
                        # elif vehicle_type == 'TRUK SEMI TRAILER':
                        #     widening = dict_13[int(radius_cell)][1] * total_lane
                        #
                        # else:
                        #     widening = None

                        return [curve_type, float(radius_cell), float(dict_2[int(speed_cell)][0]), None, None, None, None, None, \
                                None, float(dict_10[int(speed_cell)][(0)]), None, None, None, length_curve, length_curve, \
                                float(dict_3[int(speed_cell)][0]), None, None, None, None, external_distance, middle_ordinate,
                                tangent_length, None, None, '', '', float(sight_distance), \
                                float(ruwasja), float(rs), float(m), '', '', float(widening), '', '', [], [], [], [], [], [], [], [], '',turn_direction] #total 46

                else:  # ini kalau S > L
                    m = round(float(rs) * (1 - math.cos(math.radians(28.65 * sight_distance) / float(rs))) + 0.5 *(sight_distance - length_curve) * math.sin(math.radians(28.65 * sight_distance / float(rs))), 1)

                    if float(ruwasja) < m:
                        return "false:Not Enough Horizontal Sightline Offset Space (M space < M required)"
                        

                    else:
                        widening = dict_14[int(radius_cell)][int(speed_cell)]

                        return [curve_type, float(radius_cell), float(dict_2[int(speed_cell)][0]), None, None, None, None, None, \
                                None, float(dict_10[int(speed_cell)][(0)]), None, None, None, length_curve, length_curve, \
                                float(dict_3[int(speed_cell)][0]), None, None, None, None, external_distance, middle_ordinate,
                                tangent_length, None, None, '', '', float(sight_distance), \
                                float(ruwasja), float(rs), float(m), '', '', float(widening), '', '', [], [], [], [], [], [], [], [], '',
                                turn_direction]  # total 46




                    # if float(ruwasja) >= m:
                    # print("SSD            = " + str(ssd))
                    # print("PSD            = " + str(psd))
                    # print("Sight Distance = " + str(sight_distance))
                    # print("M dibutuhkan   = " + str(m))
                    # print("M dibutuhkan \u2264 M tersedia (HASIL OK)")

                    # return [sight_distance, ruwasja, rs, m]

                    # else:
                    #     return False
                    # print("M dibutuhkan > M tersedia (HASIL NOT OK)")

                # else:
                # print("Jarak Pandang > LC (HASIL NOT OK)")

            elif psd == "N":
                ssd = dict_11[int(speed_cell)][(0)]
                sight_distance = ssd

                if rs == None:
                    return "false: Rs is blank. Please check your Ms. Excel"

                elif ruwasja == None:
                    return "false: HSO is blank. Please check your Ms. Excel"

                elif sight_distance <= length_curve: # ini kalau S <= L
                    m = round(float(rs) * (1 - math.cos(math.radians(28.65 * sight_distance) / float(rs))), 1)

                    if float(ruwasja) < m:
                        return "false:Not Enough Horizontal Sightline Offset Space (M space < M required)"


                    else:
                        widening = dict_14[int(radius_cell)][int(speed_cell)]
                        # if vehicle_type == 'TRUK TUNGGAL':
                        #     widening = dict_14[int(radius_cell)][int(speed_cell)]
                        #
                        # elif vehicle_type == 'TRUK SEMI TRAILER':
                        #     widening = dict_13[int(radius_cell)][1] * total_lane
                        #
                        # else:
                        #     widening = None

                        return [curve_type, float(radius_cell), float(dict_2[int(speed_cell)][0]), None, None, None, None, None, \
                                None, float(dict_10[int(speed_cell)][(0)]), None, None, None, length_curve, length_curve, \
                                float(dict_3[int(speed_cell)][0]), None, None, None, None, external_distance, middle_ordinate,
                                tangent_length, None, None, '', '', float(sight_distance), \
                                float(ruwasja), float(rs), float(m), '', '', float(widening), '', '', [], [], [], [], [], [], [], [], '',
                                turn_direction]  # total 46

                else:  # ini kalau S > L
                    m = round(float(rs) * (1 - math.cos(math.radians(28.65 * sight_distance) / float(rs))) + 0.5 * (sight_distance - length_curve) * math.sin(math.radians(28.65 * sight_distance / float(rs))), 1)

                    if float(ruwasja) < m:
                        return "false:Not Enough Horizontal Sightline Offset Space (M space < M required)"


                    else:
                        widening = dict_14[int(radius_cell)][int(speed_cell)]

                        return [curve_type, float(radius_cell), float(dict_2[int(speed_cell)][0]), None, None, None, None, None, \
                                None, float(dict_10[int(speed_cell)][(0)]), None, None, None, length_curve, length_curve, \
                                float(dict_3[int(speed_cell)][0]), None, None, None, None, external_distance, middle_ordinate,
                                tangent_length, None, None, '', '', float(sight_distance), \
                                float(ruwasja), float(rs), float(m), '', '', float(widening), '', '', [], [], [], [], [], [], [], [], '',
                                turn_direction]  # total 46


                    # if float(ruwasja) >= m:
                    # print("SSD            = " + str(ssd))
                    # print("PSD            = " + str(psd))
                    # print("Sight Distance = " + str(sight_distance))
                    # print("M dibutuhkan   = " + str(m))
                    # print("M dibutuhkan \u2264 M tersedia (HASIL OK)")

                    # return [sight_distance, ruwasja, rs, m]

                    # else:
                    #     return False
                    # print("M dibutuhkan > M tersedia (HASIL NOT OK)")

                    # print("Jarak Pandang > LC (HASIL NOT OK)")
            else:
                return "false:Passing Sight Distance is blank. Please check on Ms. Excel"


            # return [curve_type, radius_cell, dict_2[int(speed_cell)][0], None, None, None, None, None,\
            #         None, dict_10[int(speed_cell)][(0)], None, None, None, length_curve, None,\
            #         dict_3[int(speed_cell)][0], None, None, None, None, external_distance, middle_ordinate, tangent_length, None, [], [], [], sight_distance, \
            #         ruwasja, rs, m, [], [], [], [], [], [], [], [], [], []]




        # ----------------------------------------------------------------------------------------------------------------------


        # ---------------------------------------------------- Full Circle with Superelevation ------------------------------------------------------------------
        elif int(radius_cell) < dict_1[int(speed_cell)][int(elevation_max_hor)] and int(radius_cell) > dict_2[int(speed_cell)][0] and int(radius_cell) > dict_4[int(speed_cell)][int(elevation_max_hor)]:
            # print('Pakai FC dengan Elevasi')
            curve_type = 'FC with Elevation'
            length_curve = round(2 * math.pi * (int(radius_cell)) * (float(delta_azimuth_hor)) / 360, 1)
            # print(dict_3[int(speed_cell)][0])

            if speed_cell < 40 or speed_cell > 120:
                return "false:FC speed is less than 40 km/h or Speed is more than 120 km/h"

            elif length_curve < dict_3[int(speed_cell)][0]: #minimal speed = 40
                # print('LC < Lmin !! Perhitungan salah !')
                return "false:Lc < Lmin"


            # print('LC = ' + str(length_curve))

            tangent_length = round(int(radius_cell) * math.tan(math.radians(float(delta_azimuth_hor) / 2)), 1)
            # print('T  = ' + str(tangent_length))

            external_distance = round(int(radius_cell) * (sec(math.radians(float(delta_azimuth_hor) / 2)) - 1), 1)
            # print('E  = ' + str(external_distance))

            middle_ordinate = round(int(radius_cell) * (1 - math.cos(math.radians(float(delta_azimuth_hor) / 2))), 1)
            # print('M  = ' + str(middle_ordinate))


            bw              = dict_5[int(rotated_lane)][(0)]
            delta           = dict_6[int(speed_cell)][(0)]

            if elevation_normal_hor == -2 and elevation_max_hor == 8:
                elevation_design            = dict_7[int(radius_cell)][int(speed_cell)]
                length_spiral_kelandaian    = round((float(one_lane_width) * float(rotated_lane) * elevation_design / 100 * bw) / delta, 1)

            elif elevation_normal_hor == -2 and elevation_max_hor == 6:
                elevation_design            = dict_8[int(radius_cell)][int(speed_cell)]
                length_spiral_kelandaian    = round((float(one_lane_width) * float(rotated_lane) * elevation_design / 100 * bw) / delta, 1)

            elif elevation_normal_hor == -2 and elevation_max_hor == 4:
                elevation_design            = dict_9[int(radius_cell)][int(speed_cell)]
                length_spiral_kelandaian    = round((float(one_lane_width) * float(rotated_lane) * elevation_design / 100 * bw) / delta, 1)


            # print('Elevation Design = ' + str(elevation_design))

            length_spiral_offset        = round(math.sqrt(24 * 0.2 * int(radius_cell)), 1)
            length_spiral_acceleration  = round((0.0214 * (float(speed_cell) ** 3)) / (float(radius_cell) * 1.2), 1)

            # print('\nLS Kelandaian Relatif\t = ' + str(length_spiral_kelandaian) + ' meter')
            # print('LS Offset Lateral\t\t = ' + str(length_spiral_offset) + ' meter')
            # print('LS Akselerasi Lateral\t = ' + str(length_spiral_acceleration) + ' meter')

            tangent_runout  = round(abs(float(elevation_normal_hor)/100) / (elevation_design / 100) * max(length_spiral_kelandaian, length_spiral_offset, length_spiral_acceleration), 1)
            # print('Tangent Runout\t\t\t = ' + str(tangent_runout) + ' meter')

            if speed_cell < 80:
                lrr = round((elevation_design / 100 - float(elevation_normal_hor) / 100) / (3.6 * 0.035) * float(speed_cell), 1)
            else:
                lrr = round((elevation_design / 100 - float(elevation_normal_hor) / 100) / (3.6 * 0.025) * float(speed_cell), 1)

            # print('Lrr\t\t\t\t\t\t = ' + str(lrr))

            if lrr <= max(length_spiral_kelandaian, length_spiral_offset, length_spiral_acceleration) + tangent_runout:
                # print('Lrr \u2264 Ls + TR(Lt)\t HASIL OK\n')
                length_spiral_design = max(length_spiral_kelandaian, length_spiral_offset, length_spiral_acceleration)
                ts_tr   = round(tangent_runout + tangent_length + 2/3*length_spiral_design, 1)

                if length_spiral_design < dict_10[int(speed_cell)][(0)]:
                    return "false:Ls < Ls min"


                #------ Jarak Bebas Samping -----#
                if psd == "Y":
                    ssd = dict_11[int(speed_cell)][(0)]
                    psd = dict_12[int(speed_cell)][(0)]
                    sight_distance = max(ssd, psd)

                    if rs == None:
                        return "false: Rs is blank. Please check your Ms. Excel"

                    elif ruwasja == None:
                        return "false: HSO is blank. Please check your Ms. Excel"

                    elif sight_distance <= length_curve:
                        m = round(float(rs) * (1 - math.cos(math.radians(28.65 * sight_distance) / float(rs))), 1)

                        if float(ruwasja) < m:
                            return "false:Not Enough Horizontal Sightline Offset Space (M space < M required)"

                        else:
                            widening = dict_14[int(radius_cell)][int(speed_cell)]
                            # if vehicle_type == 'TRUK TUNGGAL':
                            #     widening = dict_13[int(radius_cell)][0] * total_lane
                            #
                            # elif vehicle_type == 'TRUK SEMI TRAILER':
                            #     widening = dict_13[int(radius_cell)][1] * total_lane
                            #
                            # else:
                            #     widening = None

                            return [curve_type, float(radius_cell), float(dict_2[int(speed_cell)][0]), elevation_design,
                                    length_spiral_kelandaian,
                                    length_spiral_offset, length_spiral_acceleration, \
                                    lrr, None, float(dict_10[int(speed_cell)][(0)]), length_spiral_design, None, None,
                                    length_curve, length_curve,
                                    float(dict_3[int(speed_cell)][0]), None, None, None, \
                                    None, external_distance, middle_ordinate, tangent_length, tangent_runout, ts_tr, '',
                                    '', float(sight_distance), float(ruwasja), \
                                    float(rs), float(m), '', '', float(widening), '', '', [], [], [], [], [], [], [], [], '', turn_direction] #total 46


                    else:  # ini kalau S > L
                        m = round(float(rs) * (1 - math.cos(math.radians(28.65 * sight_distance) / float(rs))) + 0.5 * (sight_distance - length_curve) * math.sin(math.radians(28.65 * sight_distance / float(rs))), 1)

                        if float(ruwasja) < m:
                            return "false:Not Enough Horizontal sightline Offset Space (M space < M required)"

                        else:
                            widening = dict_14[int(radius_cell)][int(speed_cell)]

                            return [curve_type, float(radius_cell), float(dict_2[int(speed_cell)][0]), elevation_design,
                                    length_spiral_kelandaian,
                                    length_spiral_offset, length_spiral_acceleration, \
                                    lrr, None, float(dict_10[int(speed_cell)][(0)]), length_spiral_design, None, None,
                                    length_curve, length_curve,
                                    float(dict_3[int(speed_cell)][0]), None, None, None, \
                                    None, external_distance, middle_ordinate, tangent_length, tangent_runout, ts_tr, '',
                                    '', float(sight_distance), float(ruwasja), \
                                    float(rs), float(m), '', '', float(widening), '', '', [], [], [], [], [], [], [], [], '',
                                    turn_direction]  # total 46


                        # if float(ruwasja) >= m:
                            # print("SSD            = " + str(ssd))
                            # print("PSD            = " + str(psd))
                            # print("Sight Distance = " + str(sight_distance))
                            # print("M dibutuhkan   = " + str(m))
                            # print("M dibutuhkan \u2264 M tersedia (HASIL OK)")

                            # return [sight_distance, ruwasja, rs, m]

                        # else:
                        #     return False
                            # print("M dibutuhkan > M tersedia (HASIL NOT OK)")

                    # else:
                    # print("Jarak Pandang > LC (HASIL NOT OK)")

                elif psd == 'N':
                    ssd = dict_11[int(speed_cell)][(0)]
                    sight_distance = ssd

                    if rs == None:
                        return "false: Rs is blank. Please check your Ms. Excel"

                    elif ruwasja == None:
                        return "false: HSO is blank. Please check your Ms. Excel"

                    elif sight_distance <= length_curve:
                        m = round(float(rs) * (1 - math.cos(math.radians(28.65 * sight_distance) / float(rs))), 1)

                        if float(ruwasja) < m:
                            return "false:Not Enough Horizontal Sightline Offset Space (M space < M required)"

                        else: #------------------Widening---------------------#
                            widening = dict_14[int(radius_cell)][int(speed_cell)]
                            # if vehicle_type == 'TRUK TUNGGAL':
                            #     widening = dict_13[int(radius_cell)][0] * total_lane
                            #
                            # elif vehicle_type == 'TRUK SEMI TRAILER':
                            #     widening = dict_13[int(radius_cell)][1] * total_lane
                            #
                            # else:
                            #     widening = None

                            return [curve_type, float(radius_cell), float(dict_2[int(speed_cell)][0]), elevation_design,
                                    length_spiral_kelandaian,
                                    length_spiral_offset, length_spiral_acceleration, \
                                    lrr, None, float(dict_10[int(speed_cell)][(0)]), length_spiral_design, None, None,
                                    length_curve, length_curve,
                                    float(dict_3[int(speed_cell)][0]), None, None, None, \
                                    None, external_distance, middle_ordinate, tangent_length, tangent_runout, ts_tr, '',
                                    '', float(sight_distance), float(ruwasja), \
                                    float(rs), float(m), '', '', float(widening), '', '', [], [], [], [], [], [], [], [], '',
                                    turn_direction]  # total 46
                        # if float(ruwasja) >= m:
                            # print("SSD            = " + str(ssd))
                            # print("PSD            = " + str(psd))
                            # print("Sight Distance = " + str(sight_distance))
                            # print("M dibutuhkan   = " + str(m))
                            # print("M dibutuhkan \u2264 M tersedia (HASIL OK)")

                            # return [sight_distance, ruwasja, rs, m]

                        # else:
                        #     return False
                            # print("M dibutuhkan > M tersedia (HASIL NOT OK)")


                    else:  # ini kalau S > L
                        m = round(float(rs) * (1 - math.cos(math.radians(28.65 * sight_distance) / float(rs))) + 0.5 * (sight_distance - length_curve) * math.sin(math.radians(28.65 * sight_distance / float(rs))), 1)

                        if float(ruwasja) < m:
                            return "false:Not Enough Horizontal Sightline Offset Space (M space < M required)"


                        else:
                            widening = dict_14[int(radius_cell)][int(speed_cell)]

                            return [curve_type, float(radius_cell), float(dict_2[int(speed_cell)][0]), elevation_design,
                                    length_spiral_kelandaian,
                                    length_spiral_offset, length_spiral_acceleration, \
                                    lrr, None, float(dict_10[int(speed_cell)][(0)]), length_spiral_design, None, None,
                                    length_curve, length_curve,
                                    float(dict_3[int(speed_cell)][0]), None, None, None, \
                                    None, external_distance, middle_ordinate, tangent_length, tangent_runout, ts_tr, '',
                                    '', float(sight_distance), float(ruwasja), \
                                    float(rs), float(m), '', '', float(widening), '', '', [], [], [], [], [], [], [], [], '',
                                    turn_direction]  # total 46
                        # print("Jarak Pandang > LC (HASIL NOT OK)")


                # return [curve_type, radius_cell, dict_2[int(speed_cell)][0], elevation_design, length_spiral_kelandaian,
                #         length_spiral_offset, length_spiral_acceleration, \
                #         lrr, None, dict_10[int(speed_cell)][(0)], length_spiral_design, None, None, length_curve, None,
                #         dict_3[int(speed_cell)][0], None, None, None, \
                #         None, external_distance, middle_ordinate, tangent_length, tangent_runout, ts_tr, [], [], sight_distance, ruwasja, \
                #         rs, m, [], [], [], [], [], [], [], [], [], []]

                else:
                    return "false:Passing Sight Distance is blank. Please check on Ms. Excel"

            else:
                # print('Lrr \u003E Ls + TR(Lt)\t HASIL NOT OK\n')
                return "false:Lrr > Ls + TR(Lt)"

        else:
            return "false:Radius cannot be designed as a Full-Circle because its radius is not fulfill the conditions. Please change either the radius min & max or the curve type option on Ms. Excel"









            # ----------------------------------------------------------------------------------------------------------------------
        # def is_integer(n):
        #     try:
        #         float(n)
        #         print(str(dict_5[float(rotated_lane)][0]) + ' Tabel 5.22')
        #     except ValueError:
        #         return print(str(dict_5[int(rotated_lane)][0]) + ' Tabel 5.22')
        #     else:
        #         return float(n).is_integer()
        # is_integer(rotated_lane)