#version 130

uniform bool Initial;
uniform float Size_X;
uniform float Size_Y;
uniform float Start_X;
uniform float Start_Y;
uniform float Batch_W;
uniform float Batch_H;
uniform sampler2D AltMapTex;
uniform sampler2D PrevKMapTex;
out highp vec4 PixelK;

highp float GetData(sampler2D sampler, vec2 coord)
{
	vec2 Size_XY = vec2(Size_X, Size_Y);
	vec2 texcrd = coord + step(coord, vec2(0.0)) * Size_XY;
	texcrd -= step(Size_XY, texcrd) * Size_XY;
	// return textureLod(sampler, texcrd / vec2(Size_X - 1.0, Size_Y - 1.0), 0).r;
	return texelFetch(sampler, ivec2(floor(texcrd)), 0).r;
}

void main()
{
	if(gl_FragCoord.x >= Size_X || gl_FragCoord.y >= Size_Y) discard;
	highp float alt_cur = GetData(AltMapTex, gl_FragCoord.xy);
	highp float k_max;
	k_max = Initial ? 0.0 : GetData(PrevKMapTex, gl_FragCoord.xy);
	
	for(float y = 0.0; y < Batch_H; y += 1.0)
	{
		for(float x = 0.0; x < Batch_W; x += 1.0)
		{
			vec2 crd = vec2(Start_X, Start_Y) + vec2(x, y);
			float dist = max(length(crd), 1.0);
			highp float alt_xyv = GetData(AltMapTex, gl_FragCoord.xy + crd);
			k_max = max(k_max, (alt_xyv - alt_cur) / dist);
		}
	}
	
	PixelK = vec4(k_max);
}
