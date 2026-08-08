"""Repository-wide Openpyxl chart compatibility settings.

Python imports ``sitecustomize`` automatically at startup when this repository is on
``sys.path`` (which it is when running ``python update_model.py`` from the repo).

The equity-research dashboards intentionally keep chart helper tables in hidden
columns. Openpyxl defaults charts to ``visible_cells_only=True``, which serializes as
``<plotVisOnly val=\"1\"/>``. Microsoft Excel then suppresses every series whose source
cells are hidden, making the Analysis Charts sheet appear empty even though the chart
objects and data are present.

Set the default to False for all newly created Openpyxl charts so Excel plots hidden
helper cells while the workbook stays clean for the user.
"""

try:
    from openpyxl.chart._chart import ChartBase

    _original_chartbase_init = ChartBase.__init__

    def _chartbase_init_plot_hidden(self, *args, **kwargs):
        _original_chartbase_init(self, *args, **kwargs)
        self.visible_cells_only = False

    # Make the patch idempotent in case sitecustomize is reloaded interactively.
    if not getattr(ChartBase, "_equity_research_plot_hidden_patch", False):
        ChartBase.__init__ = _chartbase_init_plot_hidden
        ChartBase._equity_research_plot_hidden_patch = True
except Exception:
    # Never block the model updater because of a presentation-only compatibility patch.
    pass
