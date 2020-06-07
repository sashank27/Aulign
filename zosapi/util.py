# Function to extract Zernike coefficients from the results file 
# obtained from the OpticStudio, using regex parsing

def extractZernikeCoefficents(file):
    import codecs
    import re

    pattern = "^Z\s*(\d*)\s*([-]?[0-9]+[,.]?[0-9]*).*$"

    x = []
    for line in codecs.open(file, 'r', encoding='utf16'):
        match = re.search(pattern, line)
        if match is not None:
            x.append(float(match.group(2)))

    coefficients = np.array(x)
    return coefficients