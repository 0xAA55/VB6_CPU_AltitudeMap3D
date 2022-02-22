attribute vec2 iPosition;

void main()
{
    gl_Position = vec4(iPosition, 0.0, 1.0);
}