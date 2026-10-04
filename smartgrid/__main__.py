"""Explicit startup: importing SmartGrid never moves windows or opens the UI."""
from __future__ import annotations
import argparse
from dataclasses import asdict
import json
import logging
from logging.handlers import RotatingFileHandler
import os
from pathlib import Path
import sys
import tempfile
import time
from . import __version__


def demo_backend():
    from .platform.fake import FakeBackend
    from .core.models import ApplicationRef, Display, Rect, WindowRef, WindowRecord
    displays=[Display('demo-primary','Demo display',Rect(0,0,1920,1080),primary=True)]
    apps=[ApplicationRef(key,name) for key,name in [('browser','Browser'),('editor','Editor'),('terminal','Terminal'),('files','Files')]]
    windows=[WindowRecord(WindowRef(i+1,i+1,f'demo-{i}'),a.id,a.name,Rect(80+i*60,60+i*40,700,500),'demo-primary') for i,a in enumerate(apps)]
    return FakeBackend(displays,windows,apps)


def report(message,icon=0x10):
    """Print to the console; under pythonw (desktop shortcut) show a Windows dialog instead."""
    if sys.stderr is not None:
        print(message,file=sys.stderr)
    elif os.name=='nt':
        import ctypes
        ctypes.windll.user32.MessageBoxW(None,message,'SmartGrid',icon)


def main(argv=None):
    parser=argparse.ArgumentParser(description='SmartGrid · Windows layout studio')
    parser.add_argument('--version',action='version',version=__version__)
    parser.add_argument('--demo',action='store_true',help='Simulated desktop; real windows are never moved')
    parser.add_argument('--diagnose',action='store_true',help='Read-only JSON diagnostics; titles and paths omitted')
    parser.add_argument('--data-dir',type=Path,help='Alternative settings folder')
    args=parser.parse_args(argv)
    if not args.demo and os.name!='nt':
        parser.error('Native placement requires Windows. Use --demo to explore the Studio.')
    from .storage.repository import Repository,default_directory
    from .core.controller import Controller
    temporary=tempfile.TemporaryDirectory(prefix='smartgrid-demo-') if args.demo and args.data_dir is None else None
    directory=args.data_dir or (Path(temporary.name) if temporary else default_directory())
    guard=None
    controller=None
    try:
        discovery_started=time.perf_counter()
        if args.demo:
            backend=demo_backend()
        else:
            from .platform.windows.instance import InstanceGuard
            if not args.diagnose:
                guard=InstanceGuard()
                if not guard.acquired:
                    report('SmartGrid is already running.\n\nIts icon is near the clock; on Windows 11, '
                           'click the ^ arrow on the taskbar to show it.',0x40)
                    return 1
            from .platform.windows.backend import WindowsBackend
            backend=WindowsBackend()
        controller=Controller(backend,Repository(directory,read_only=args.diagnose))
        if args.diagnose:
            print(json.dumps({'version':__version__,'platform':sys.platform,'capabilities':backend.capabilities,
                'discovery_seconds':round(time.perf_counter()-discovery_started,3),
                'eligible_windows':sum(w.eligible for w in controller.windows),
                'desktop_scope':{'id':getattr(backend,'desktop_id',None),'identified':getattr(backend,'desktop_known',True)},
                'displays':[asdict(d) for d in controller.displays],
                'windows':[{'id':i,'eligible':w.eligible,'reason':w.exclusion_reason,'state':w.state,'display_id':w.display_id,'rect':asdict(w.rect)} for i,w in enumerate(controller.windows)],
                'configuration_errors':controller.repository.errors},ensure_ascii=False,indent=2))
            return 0
        directory.mkdir(parents=True,exist_ok=True)
        handler=RotatingFileHandler(directory/'smartgrid.log',maxBytes=1024*1024,backupCount=3,encoding='utf-8')
        handler.setFormatter(logging.Formatter('%(asctime)s %(levelname)s %(name)s: %(message)s'))
        logging.getLogger('smartgrid').addHandler(handler)
        logging.getLogger('smartgrid').setLevel(logging.INFO)
        from .ui.app import run_ui
        code=run_ui(controller,[sys.argv[0]])
        if args.demo:
            return code
        # A native thread blocked by a hung application must not keep the
        # process (and the single-instance lock) alive after Quit.
        controller.quit()
        if guard: guard.close()
        logging.shutdown()
        os._exit(code)
    except ImportError as error:
        report(f'Missing dependency: {error}. Run smartgrid.bat to repair the installation.')
        return 2
    except Exception as error:
        logging.getLogger('smartgrid').exception('Startup failed')
        report(f'SmartGrid could not start: {error}')
        return 2
    finally:
        if controller is not None and not args.diagnose:
            controller.quit()
        if guard:
            guard.close()
        if temporary:
            temporary.cleanup()


if __name__=='__main__':
    raise SystemExit(main())
