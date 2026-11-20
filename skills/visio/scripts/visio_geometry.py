"""Native drawing geometry independent of Visio and Windows."""
import math


def cylinder_paths(rear, front, radius, end_ratio, rear_scale=1):
    """Closed cubic paths for a cylinder, with optional apparent end-size difference."""
    dx, dy = front[0] - rear[0], front[1] - rear[1]
    length = math.hypot(dx, dy)
    if not length or radius <= 0 or not 0 < end_ratio <= 1 or rear_scale <= 0:
        raise ValueError('Cylinder needs distinct end centers, positive radius/scale and end_ratio in (0, 1]')
    ux, uy = dx / length, dy / length
    axial_radius = radius * end_ratio
    # Scale the projected ellipse to a unit circle. The shared tangent
    # normals have this axial component by equality of their support lines.
    tangent_x = (rear_scale - 1) * axial_radius / length
    if not abs(tangent_x) < 1:
        raise ValueError('Cylinder end contours are nested; no shared side tangents')
    theta = math.acos(tangent_x)
    center_delta = length / axial_radius

    def arc(cx, r, start, stop):
        count = math.ceil(abs(stop-start)/(math.pi/2))
        step = (stop-start)/count
        k = 4/3 * math.tan(step/4)
        points = [(cx+r*math.cos(start), r*math.sin(start))]
        for i in range(count):
            a,b = start+i*step, start+(i+1)*step
            x1,y1,x2,y2 = math.cos(a),math.sin(a),math.cos(b),math.sin(b)
            points += [(cx+r*(x1-k*y1),r*(y1+k*x1)),
                       (cx+r*(x2+k*y2),r*(y2-k*x2)),
                       (cx+r*x2,r*y2)]
        return points

    def side(p,q):
        return [tuple(a+(b-a)*f for a,b in zip(p,q)) for f in (1/3,2/3)] + [q]

    near = arc(center_delta,1,theta,-theta)
    far = arc(0,rear_scale,-theta,theta-2*math.pi)
    # The side shape extends under the cap, avoiding rasterization gaps.
    body = [far[-1]] + side(far[-1],near[0]) + near[1:] + side(near[-1],far[0]) + far[1:]
    cap = near + arc(center_delta,1,-theta,theta-2*math.pi)[1:]
    cap[-1] = cap[0]

    def on_page(points):
        return tuple(v for x, y in points for v in (
            rear[0] + x * axial_radius * ux - y * radius * uy,
            rear[1] + x * axial_radius * uy + y * radius * ux))

    return on_page(body), on_page(cap)
