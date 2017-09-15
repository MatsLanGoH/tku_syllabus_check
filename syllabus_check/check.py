# encoding: utf-8


class Check:
    """
    A Check object should contain these elements:
    msg: A string describing the desired result of the check.
    fun: The test function to conduct the test 
         (this should return an error message if test fails)
    args: (Optional) arguments to pass into the test function.

    After instantation, the passes() method is used to determine test result.
    """
    def __init__(self, msg, fun, *args):
        self.msg = msg
        self.fun = fun
        self.my_args = args
        self.errors = ""

    def passes(self):
        # Returns False if an error message has been created.
        self.errors = self.fun(*self.my_args)
        return len(self.errors) == 0

