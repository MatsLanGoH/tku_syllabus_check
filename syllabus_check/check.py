# encoding: utf-8


class Check:
    def __init__(self, msg, fun, *args):
        self.msg = msg
        self.fun = fun
        self.my_args = args

    def passes(self):
        return self.fun(*self.my_args)

