import traceback
from Gui import *
if __name__ == "__main__":
    try:
        Application.EnableVisualStyles()
        form = BeamForm()
        Application.Run(form)

    except Exception:
        traceback.print_exc()
