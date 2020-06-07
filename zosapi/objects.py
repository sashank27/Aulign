"""
Object definitions
"""

class Parameter(object):

    def __init__(self, name, optimal, minval, maxval, step, optimal1=None):
        self.name = name
        self.optimal = optimal
        self.minval = minval
        self.maxval = maxval
        self.step = step
        
        if optimal1 is not None:
            self.optimal1 = optimal1

class Component(object):
    
    def __init__(self, surface, revSurface, compName):
        self.surface = surface
        self.compName = compName
        if revSurface is not None:
            self.revSurface = revSurface
    
    def createParameter(self, name, optimal, minval, maxval, step, optimal1=None):
        
        param = Parameter(name, optimal, minval, maxval, step, optimal1)
        setattr(self, name, param)
        print(name, param)
