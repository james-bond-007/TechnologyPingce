# -*- coding: utf-8 -*-
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

# for chinese show
plt.rcParams['axes.unicode_minus'] = False
myfont = fm.FontProperties(fname='SimHei.ttf')

"""
bar 
"""


def plt_bar(names, values, picture):
    plt.figure(figsize=(9, 3))
    plt.bar(names, values)  # bar
    plt.suptitle('柱状图', fontproperties=myfont)
    plt.savefig(picture)  # eps, jpeg, jpg, pdf, pgf, png, ps, raw, rgba, svg, svgz, tif, tiff 都可以

    # plt.show()


"""
line
"""


def plt_plot(names, values, picture):
    plt.figure(figsize=(9, 3))

    plt.plot(names, values)  # line
    plt.suptitle('折线图', fontproperties=myfont)
    plt.savefig(picture)  # eps, jpeg, jpg, pdf, pgf, png, ps, raw, rgba, svg, svgz, tif, tiff 都可以
    # plt.show()


"""
pie
"""


def plt_pie(labels, sizes, picture):
    # Pie chart, where the slices will be ordered and plotted counter-clockwise:
    explode = (0, 0.1, 0, 0)  # only "explode" the 2nd slice (i.e. 'Hogs')

    fig1, ax1 = plt.subplots()
    ax1.pie(sizes, explode=explode, labels=labels, autopct='%1.1f%%',
            shadow=True, startangle=90)
    ax1.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
    plt.suptitle('饼状图', fontproperties=myfont)
    plt.savefig(picture)  # eps, jpeg, jpg, pdf, pgf, png, ps, raw, rgba, svg, svgz, tif, tiff 都可以
    # plt.show()


"""
scatter
"""


def plt_scatter(names, values, picture):
    plt.figure(figsize=(9, 3))

    plt.scatter(names, values)  # scatter
    plt.suptitle('散点图', fontproperties=myfont)
    plt.savefig(picture)  # eps, jpeg, jpg, pdf, pgf, png, ps, raw, rgba, svg, svgz, tif, tiff 都可以
    # plt.show()


if __name__ == '__main__':
    names = ['group_a', 'group_b', 'group_c']
    values = [1, 10, 100]

    labels = 'Frogs', 'Hogs', 'Dogs', 'Logs'
    sizes = [15, 30, 45, 10]

    plt_bar(names, values, "柱状图.jpg")
    plt_plot(names, values, "折线图.jpg")
    plt_pie(labels, sizes, "饼状图.jpg")
    plt_scatter(names, values, "散点图.jpg")
