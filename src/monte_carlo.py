import argparse
import uno


def main():
    # get the uno component context from the PyUNO runtime
    localContext = uno.getComponentContext()

    # create the UnoUrlResolver
    resolver = localContext.ServiceManager.createInstanceWithContext(
                    "com.sun.star.bridge.UnoUrlResolver", localContext )

    # connect to the running office
    ctx = resolver.resolve( "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" )
    smgr = ctx.ServiceManager

    # get the central desktop object
    desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)

    # access the current calc document
    model = desktop.getCurrentComponent()

    parser = argparse.ArgumentParser(description='Runs a Monte Carlo simulation on a spreadsheet')
    parser.add_argument('--model_sheet', dest='model_sheet', help='Model sheet name', required=True)
    parser.add_argument('--output_cell', dest='output_cell', help='Output cell in model sheet', required=True)
    parser.add_argument('--variables_sheet', dest='variables_sheet', help='Variables sheet name', required=True)
    parser.add_argument('--n_sims', dest='n_sims', help='The number of simulations to run', type=int, default=10000)
    args = parser.parse_args()

    # get the risk variables
    variable_range = model.getSheets().getByName(args.variables_sheet).getCellRangeByName("A2:C101")
    NAME_COLUMN = 0
    MIN_VALUE_COLUMN = 1
    MAX_VALUE_COLUMN = 2
    OUTPUT_COLUMN = 3
    risk_variables = []
    for row in variable_range.getRows():
        # Indexing is [column, row] (not [r, c])
        name = row.getCellByPosition(NAME_COLUMN, 0).getString()
        if not name:
            break
        risk_variables.append({
            'name': name,
            'min': row.getCellByPosition(MIN_VALUE_COLUMN, 0).getValue(),
            'max': row.getCellByPosition(MAX_VALUE_COLUMN, 0).getValue(),
            'output_cell': row.getCellByPosition(OUTPUT_COLUMN, 0)
        })

    # evaluate the model with random values
    import numpy
    model_output_cell = model_output = model.getSheets().getByName(args.model_sheet).getCellRangeByName(args.output_cell)
    output_values = []
    for i in range(0, args.n_sims):
        for var in risk_variables:
            output_value = numpy.random.uniform(var['min'], var['max'])
            var['output_cell'].setValue(output_value)
        output_values.append(model_output_cell.getValue())

    # draw the histogram output
    print("Variance: {}".format(numpy.var(output_values, axis=0)))

    from matplotlib import pyplot as plt
    plt.hist(output_values, bins=16)
    plt.show()

    # Do a nasty thing before exiting the python process. In case the
    # last call is a oneway call (e.g. see idl-spec of insertString),
    # it must be forced out of the remote-bridge caches before python
    # exits the process. Otherwise, the oneway call may or may not reach
    # the target object.
    # I do this here by calling a cheap synchronous call (getPropertyValue).
    ctx.ServiceManager

if __name__ == '__main__':
    main()