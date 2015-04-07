

# natlink13 library
from natlinkmain import (setCheckForGrammarChanges)
from natlinkutils import (GrammarBase)

# standard library
import time
from urllib2 import urlopen


# interpolated from "H"askell

H_RULES  = '''

<dgndictation> imported;
<dgnwords> imported;
<dgnletters> imported;

<test> exported  = ({Transport})+;

'''
H_LISTS  = {"Transport": ["foot","bike","public transit","car"]}
H_EXPORT = "test"

H_SERVER_HOST = '192.168.56.1'
H_SERVER_PORT = 8080

server_address = "http://%s:%s/" % (H_SERVER_HOST, H_SERVER_PORT)

# H_CLIENT_HOST = '192.168.56.101'
# H_CLIENT_PORT = 8080



# the grammar

class NarcissisticGrammar(GrammarBase):
    ''' 'Narcissistic' because:

    * load(.., allResults=1)     means: every recognition triggers gotResultsObject
    * load(.., hypothesis=1)     means: every hypothesis, before the recognition, triggers gotHypothesis
    * activate(.., exclusive=1)  means: deactivate every other non-exclusive rule

    (when both flags are set, NarcissisticGrammar.gotResultsObject is called on
    every recognition of every exclusive rule, including this class's rules
    of course, though I only expect this class to be active).

    '''

    gramSpec = H_RULES

    def initialize(self):
        self.set_rules(H_RULES, [H_EXPORT])
        self.set_lists(H_LISTS)
        self.set_exports(["test"]) # should be idempotent

    # called when speech is detected before recognition begins.
    def gotBegin(self, moduleInfo ):
        # handleDGNContextResponse(timeit("/context", urlopen, ("%s/context" % server_address), timeout=0.1))
        # TODO parameterize " Context"

        print
        print "-  -  -  -  gotBegin  -  -  -  -"
        print moduleInfo

    def gotHypothesis(self, words):
        print
        print "-  -  -  -  gotHypothesis  -  -  -  -"
        print words

    def gotResultsInit(self, words, results):
        print
        print "-  -  -  -  gotResultsInit  -  -  -  -"
        print results

    def gotResultsObject(self, recognitionType, resultsObject):
        words = next(get_results(resultsObject), [])
        text  = munge_and_flatten(words)
        url   = "%s/recognition/%s" % (server_address, text)
        # TODO parameterize "recognition"
        self.activateAll(["test"]) # should be idempotent

        try:
            response = timeit("/recognition/...", urlopen, url, timeout=0.1)
            handleDGNSerializedResponse(self, response)
        except Exception as e:
            print "sending the request and/or handling the response threw:"
            print e

        # don't print until the request is sent the response is handled
        try:
            print
            print "---------- gotResultsObject ----------"
            print "words  =", words
            print "status =", response.getcode()
            print "body   =", list(response)
        except NameError:
            pass

    # TODO    must it reload the grammar?
    # TODO    should include export for safety?
    def set_rules(self, rules, exports):
        self.gramSpec = rules
        self.load(rules, allResults=1, hypothesis=1)
        self.set_exports(exports)

    # TODO must it reload the grammar?
    def set_lists(self, lists):
        for (lhs, rhs) in lists.items():
            self.setList(lhs, rhs)

    def set_exports(self, exports):
        self.activateSet(exports, exclusive=1)


# API

def handleDGNSerializedResponse(self, response):
    if response.getcode() != 200:
        return

    j = list(response)[0]
    o = DGNSerializedResponse.fromJSON(j)

    # TODO does order matter? Before/after loading?
    if o is not None:
        if o.dgnSerializedRules is not None and o.dgnSerializedExports is not None:
            self.set_rules(o.dgnSerializedRules, o.dgnExport)
        if o.dgnSerializedLists is not None:
            self.set_lists(o.dgnSerializedLists)
        if o.dgnSerializedRules is None and o.dgnSerializedExports is not None: # just switch context
            self.set_exports(o.dgnSerializedExports)

class DGNSerializedResponse(object):

    @classmethod
    def fromJSON(cls, j):
        try:
            d = json.load(j)
            dgnSerializedRules = d["dgnSerializedRules"]
            dgnSerializedExports = d["dgnSerializedExports"]
            dgnSerializedLists = d["dgnSerializedLists"]
            o = cls(dgnSerializedRules, dgnSerializedExports, dgnSerializedLists)
            return o
        except Exception:
            return None

    # TODO something dynamic using __dict__ and *args or something:
    # generated code is less readable/debuggable, but the generating code is simpler
    def __init__(self, dgnSerializedRules, dgnSerializedExports, dgnSerializedLists):
        self.dgnSerializedRules = dgnSerializedRules
        self.dgnSerializedExports = dgnSerializedExports
        self.dgnSerializedLists = dgnSerializedLists

    # def toJSON(self):



# helpers

# current time in milliseconds
def now():
    return int(time.clock() * 1000)

def first_result(resultsObject):
    return next(get_results(resultsObject), None)

def get_results(resultsObject):
    '''iterators are more idiomatic'''
    try:
        for number in xrange(10):
            yield resultsObject.getWords(number)
    except:
        return

def munge_and_flatten(words):
    '''
    >>> munge_and_flatten(['spell', r'a\\spelling-letter\\A', r',\\comma\\comma', r'a\\determiner', 'letter'])
    'spell A , a letter'
    '''
    return ' '.join(word.split(r'\\')[0] for word in words)

# http://stackoverflow.com/questions/1685221/accurately-measure-time-python-function-takes
def timeit(message, callback, *args, **kwargs):
    before = time.clock()
    result = callback(*args,**kwargs)
    after = time.clock()
    print message, ': ', (after - before) * 1000, 'ms'
    return result




# boilerplate

GRAMMAR = None # mutable global

def load():
    global GRAMMAR
    setCheckForGrammarChanges(1) # automatically reload on file change (not only when microphone toggles on)
    GRAMMAR = NarcissisticGrammar()
    GRAMMAR.initialize()

def unload():
    global GRAMMAR
    if GRAMMAR:
        GRAMMAR.unload()
    GRAMMAR = None

load()


