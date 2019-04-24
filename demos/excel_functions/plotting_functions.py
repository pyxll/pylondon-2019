"""
Functions for plotting DataFrames using seaborn
"""
from pyxll import xl_func, xl_app, xlfCaller
from matplotlib.backends.backend_agg import FigureCanvasAgg as FigureCanvas
from matplotlib.figure import Figure
import seaborn as sb
import os


@xl_func("dataframe, str, bool kde, bool norm, str", macro=True)
def plot_dist(data, column, kde=False, norm=False, name=None):
    fig = Figure(figsize=(8, 6), dpi=75, facecolor=(1, 1, 1), edgecolor=(0, 0, 0))
    ax = fig.add_subplot(111)

    if name is None:
        name = f"{column} distribution"

    fit = None
    if norm:
        from scipy.stats import norm as fit

    # plot using seaborn
    sb.distplot(data[column].dropna(), ax=ax, fit=fit, kde=kde)

    # plot the figure in Excel
    return plot_to_excel(fig, name)


@xl_func("dataframe, str, str, str, str: var", macro=True)
def plot_scatter(data, x, y, hue=None, name=None):
    """Creates a scatter plot where x and y are columns in data"""
    # create the axis to plot to
    fig = Figure(figsize=(8, 6), dpi=75, facecolor=(1, 1, 1), edgecolor=(0, 0, 0))
    ax = fig.add_subplot(111)

    if name is None:
        name = f"{x} vs. {y}"

    # allow using a named index as a data column
    data = data.reset_index()

    # do the scatter plot using seaborn
    sb.scatterplot(x=x, y=y, hue=hue, data=data, ax=ax)

    # plot the figure in Excel
    return plot_to_excel(fig, name)


@xl_func("dataframe, str, var, str, int[]: var", macro=True)
def plot_area(data, name, stacked=True, legend_loc="upper right", ylim=None):
    # create the axis to plot to
    fig = Figure(figsize=(8, 6), dpi=75, facecolor=(1, 1, 1), edgecolor=(0, 0, 0))
    ax = fig.add_subplot(111)

    # Do the area plot
    data.plot(kind='area', stacked=stacked, ax=ax)
    ax.legend(loc=legend_loc)
    if ylim:
        ax.set_ylim(0, 1)

    # plot the figure in Excel
    return plot_to_excel(fig, name)


def plot_to_excel(figure, name):
    """Plot a figure in Excel"""
    # write the figure to a temporary image file
    filename = os.path.join(os.environ["TEMP"], "xlplot_%s.png" % name)
    canvas = FigureCanvas(figure)
    canvas.draw()
    canvas.print_png(filename)

    # Show the figure in Excel as a Picture object on the same sheet
    # the function is being called from.
    xl = xl_app()
    caller = xlfCaller()
    sheet = xl.Range(caller.address).Worksheet

    # if a picture with the same figname already exists then get the position
    # and size from the old picture and delete it.
    for old_picture in sheet.Pictures():
        if old_picture.Name == name:
            height = old_picture.Height
            width = old_picture.Width
            top = old_picture.Top
            left = old_picture.Left
            old_picture.Delete()
            break
    else:
        # otherwise place the picture below the calling cell.
        top_left = sheet.Cells(caller.rect.last_row + 2, caller.rect.last_col + 1)
        top = top_left.Top
        left = top_left.Left
        width, height = figure.bbox.bounds[2:]

    # insert the picture
    # Ref: http://msdn.microsoft.com/en-us/library/office/ff198302%28v=office.15%29.aspx
    picture = sheet.Shapes.AddPicture(Filename=filename,
                                      LinkToFile=0,  # msoFalse
                                      SaveWithDocument=-1,  # msoTrue
                                      Left=left,
                                      Top=top,
                                      Width=width,
                                      Height=height)

    # set the name of the new picture so we can find it next time
    picture.Name = name

    # delete the temporary file
    os.unlink(filename)

    return "[Plotted '%s']" % name
