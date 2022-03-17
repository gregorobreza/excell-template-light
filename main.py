import math
import xlsxwriter
import argparse


def calculation(width: float, length: float):
    p1 = round(0.2 * (5 ** math.log10(width)), 1)
    N1 = int(round((width / p1)))

    p2 = round(0.2 * (5 ** math.log10(length)), 1)
    N2 = int(round(length / p2))

    return p1, N1, p2, N2


def get_values(p: float, N: int, final: float):
    p_n = round(p/2, 1)
    points = [0, p_n]

    for i in range(N-1):
        p_n += p
        points.append(round(p_n, 1))

    if points[-1] == final:
        return points
    elif points[-1] > final:
        del points[-1]
        return points

    zadnja = round(points[-1]+round(p/2, 1), 1)
    points.append(zadnja)
    if points[-1] < final:
        points.append(final)
    elif points[-1] > final:
        points[-1] = final

    return points


def create_excel(title: str, points_w: list, points_l: list):
    workbook = xlsxwriter.Workbook(f'{title}.xlsx')
    worksheet = workbook.add_worksheet()

    cell_format = workbook.add_format(
        {'align': "center", "bg_color": "yellow", "border": 1, "bold": True})
    cell_format2 = workbook.add_format({'align': "center", "border": 1})

    worksheet.write(1, 0, "Razdalja [m]", cell_format2)
    worksheet.write(2, 1, "Točka", cell_format)
    for i, j in enumerate(points_w):
        worksheet.write(i+3, 0, j, cell_format2)
        worksheet.write(i+3, 1, i+1, cell_format)

    for i, j in enumerate(points_l):
        worksheet.write(1, i+2, j, cell_format2)
        worksheet.write(2, i+2, i+1, cell_format)

    for i in range(len(points_w)):
        for j in range(len(points_l)):
            worksheet.write(3+i, 2+j, "", cell_format2)

    workbook.close()


def main(title: str, width: float, length: float):
    p1, N1, p2, N2 = calculation(width, length)
    points_w = get_values(p1, N1, width)
    points_l = get_values(p2, N2, length)
    print(points_w, points_l)
    create_excel(title, points_w, points_l)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Create excell template')
    parser.add_argument("-n", '--name', type=str,
                        help='Name of excell file', required=True)
    parser.add_argument("-w", '--width', type=float,
                        help='Width of field', required=True)
    parser.add_argument("-l", '--length', type=float,
                        help='Length of field', required=True)

    args = parser.parse_args()
    main(args.name, args.width, args.length)
