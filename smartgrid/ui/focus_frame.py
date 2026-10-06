"""Nonactivating transparent focus outline drawn around the focused window."""
import math
from PySide6.QtCore import Qt, QRectF, QPropertyAnimation, QEasingCurve
from PySide6.QtGui import QPainter, QColor, QFont, QPen
from PySide6.QtWidgets import QWidget

# Windows 11 rounds top-level window corners by 8 logical pixels.
WINDOW_RADIUS = 8
HALO = 10
BADGE = 26
SYMBOLS = {'left': '←', 'right': '→', 'up': '↑', 'down': '↓'}


class FocusFrame(QWidget):
    def __init__(self, topmost=False):
        super().__init__()
        # The focus outline is not topmost: it is stacked right above the focused
        # window (see DesktopUI). Resize previews stay above everything.
        flags = (Qt.WindowType.Tool | Qt.WindowType.FramelessWindowHint |
                 Qt.WindowType.WindowDoesNotAcceptFocus | Qt.WindowType.WindowTransparentForInput)
        self.setWindowFlags(flags | Qt.WindowType.WindowStaysOnTopHint if topmost else flags)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)
        self.color='#8ce8c3'
        self.width_px=2
        self.style='outline'
        self.emphasized=False
        # Swap mode: [(direction, primary, x, y)] badges in global logical
        # coordinates, centred on the window edges.
        self.arrows=[]
        self.radius=WINDOW_RADIUS
        self.window=None
        self.setAccessibleName('Focused window outline')
        self.motion=QPropertyAnimation(self,b'geometry',self)
        self.fade=QPropertyAnimation(self,b'windowOpacity',self)
        self.target=None

    def outset(self):
        """Stroke, swap band, halo and swap badges are drawn outside the window."""
        return max(self.ring_outset(),BADGE/2+1 if self.arrows else 0)

    def ring_outset(self):
        return self.stroke()+(4 if self.emphasized else 0)+(HALO if self.style in ('glow','halo') else 0)

    def stroke(self):
        width=max(1,min(6,self.width_px))
        return max(3,width+1) if self.emphasized else width

    def show_for(self,window,rect,duration,restack=None):
        """Place around a logical window rectangle; fade in on a newly focused window.

        A newly shown Qt window starts at the top of the Z order. It is shown fully
        transparent and moved under the windows above its target (restack) before
        it becomes visible, so it never flashes over Studio or Preferences.
        """
        extra=math.ceil(self.outset())
        self.setGeometry(rect.adjusted(-extra,-extra,extra,extra))
        resting=1.0 if self.emphasized else 205/255
        if window!=self.window or not self.isVisible():
            self.window=window
            self.fade.stop()
            self.setWindowOpacity(0.0);self.show()
            if restack: restack()
            if duration:
                self.fade.setDuration(duration);self.fade.setStartValue(0.0);self.fade.setEndValue(resting);self.fade.start()
            else:
                self.setWindowOpacity(resting)
        else:
            if restack: restack()
            if self.fade.state()!=QPropertyAnimation.State.Running:
                self.setWindowOpacity(resting)
        self.update()

    def hideEvent(self,event):
        self.window=None
        super().hideEvent(event)

    def move_preview(self,rect,settings):
        if self.target==rect: return
        self.target=rect;self.motion.stop()
        duration=settings.visual_duration(60)
        if not self.isVisible() or not duration:
            self.setGeometry(rect);return
        self.motion.setDuration(duration);self.motion.setStartValue(self.geometry());self.motion.setEndValue(rect)
        self.motion.setEasingCurve({'ease-out':QEasingCurve.Type.OutCubic,'linear':QEasingCurve.Type.Linear,'ease-in-out':QEasingCurve.Type.InOutCubic,'spring':QEasingCurve.Type.OutBack}[settings.animation_curve])
        self.motion.start()

    def paintEvent(self,event):
        painter=QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        color=QColor(self.color)
        stroke=self.stroke()
        band=4 if self.emphasized else 0
        halo=HALO if self.style in ('glow','halo') else 0
        pad=math.ceil(self.outset())-self.ring_outset()
        outer=QRectF(self.rect()).adjusted(pad,pad,-pad,-pad)
        def ring(distance,width,alpha):
            # distance: from the outer edge of the widget to the centre of the ring.
            c=QColor(color);c.setAlphaF(alpha)
            painter.setPen(QPen(c,width));painter.setBrush(Qt.BrushStyle.NoBrush)
            inset=distance
            radius=max(0,self.radius+(halo+band+stroke)-distance)
            painter.drawRoundedRect(outer.adjusted(inset,inset,-inset,-inset),radius,radius)
        if halo:
            step=halo/20
            for i in range(20):
                ring((i+.5)*step,step,.55*math.exp(-((19-i)/8)**2))
        if band:
            ring(halo+band/2,band,.3)
        ring(halo+band+stroke/2,stroke,1.0)
        # Filled badge: the neighbour the arrow key swaps with; outlined: other neighbours.
        font=QFont('Segoe UI',11);font.setBold(True);painter.setFont(font)
        for direction,primary,x,y in self.arrows:
            badge=QRectF(x-self.x()-BADGE/2,y-self.y()-BADGE/2,BADGE,BADGE)
            painter.setPen(QPen(color,1));painter.setBrush(color if primary else QColor(24,30,37,240))
            painter.drawRoundedRect(badge.adjusted(.5,.5,-.5,-.5),BADGE/2,BADGE/2)
            painter.setPen(QColor('#ffffff') if primary else color)
            painter.drawText(badge,Qt.AlignmentFlag.AlignCenter,SYMBOLS[direction])
