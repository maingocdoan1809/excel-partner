# -*- coding: utf-8 -*-
"""Base of the project's page objects (pages/<screen>.py). A page object hides selectors and declares the RISK of every action:
    class ItemsPage(BasePage):
        PATH = "/items"
        def create(self, name):
            guard.check(guard.WRITE_TEST, name, "create item")      # refuses names without the test prefix
            self.t.click("role=button[name='New item']", "Click 「New item」")
            self.t.fill("input[name=name]", name)
            self.t.click("role=button[name='Save']", "Click 「Save」")
            guard.track("item", name)                               # cleanup will remove it
Every interaction goes through `t` (TC) so each action gets its own boxed screenshot."""
from qakit.core import ui


class BasePage:
    PATH = "/"

    def __init__(self, t):
        self.t = t
        self.pg = t.pg

    def open(self, path=None):
        ui.goto(self.pg, path or self.PATH)
        return self
