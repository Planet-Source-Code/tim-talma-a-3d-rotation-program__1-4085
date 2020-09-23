Drawing a 3D Object

By: Tim Talma
    ttalma@stny.rr.com

First let me start by saying that there are many faster ways to draw and
transform a 3D object. The goal of this project is to help you learn the
basics of 3D animation and what is invloved. The methods in this project
are all standard routines that many 3D drawing programs use but with more
refined algorithms, and many lookup tables. This method will work for 
simple programs but for games etc. you will probably want to use somthing
like DirectX, or openGL.


I will skip an introducing you to the concept of 3D Drawing and rotation
because there are plenty of these on the internet.

How this program works

1. First a box is created using the center of the box as the origin (0,0,0)
This box could have been created in a different way I can think of many
better ones but I'll leave that up to you to implement.

2. Next the number of points for each side is stored, so we know when to 
stop drawing edges.

3. The number of sides is saved so we know how many sides make the object.

4. We find the normal for each object. The normals are import because they
tell us what direction the plane is facing. I also reduced the normals to 
have a length of one, to aid in shading, which I'll discuss later. To find
out exactly how normals relate to plane lookup information on plane 
equations.

5. Create lookup tables. For this example I only created lookup tables for
the sin and cos functions because these are by far the longest calculations
in the project.

6. Now comes the meat of the whole drawing and rotating the object. Right
now we know where the object is floating in space about the origin, and we
are looking at it from one viewpoint. so we need to determine what we will
see when we rotate it. If we rotate it 0 degrees we will only see the front.
but if rotate it we will see more than one side. So determine the rotation
angle and which orign we want to rotate around. the formulas are as follows:

About X:
	NewX = OldX
	NewY = OldY * Cos(Angle) - OldZ * Sin(Angle)
	NewZ = OldZ * Cos(Angle) + OldY * Sin(Angle)

About Y:
	NewX = OldX * Cos(Angle) + OldZ * Sin(Angle)
	NewY = OldY
	NewZ = OldZ * Cos(Angle) - OldX * Sin(Angle)

About Z:
	NewX = OldX * Cos(Angle) + OldY * Sin(Angle)
	NewY = OldY * Cos(Angle) - OldX * Sin(Angle) 
	NewZ = OldZ

After we calculate the new positions for the points we need to rotate the
normal. We rotate the normal because this is faster than recalculating the
normal for the new plane position. 
	Next we determine if the plane is visible or not. This is done 
with the normals. You multiple the normal with the position of the view 
(commonly refered to as camera). If this value is greater than 0 the side 
is visible. 
	Then we find the color the side should be. This is again done with 
the normals. First pick the light's position. Then find the 1's complement 
vector (the vector with length of 1). and put these into the following 
equation:
	ambient * Max Light Value * (Plane normal * Lights 1's compliment)

The ambient value is how bright the object is even when there is not that
much light. The Max Light Value is the brightest the light can be. I took
only one value for the Normal depending on the view because I positioned
the light at the same point as the view. So the 1's compliment would be
(0,1,0) for the top view. and the normal * the 1's compliment would result
in (0,normal.Y,0). This was done just to speed up the whole process.
	Because the normal will go from 0 to 1 the closer to parallel to 
the light source the surface becomes it will  go from darker to brighter.
Giving the apperance of shading. 
	Next choose the fill style for the object we are going to draw.
To give the object a solid look choose solid, then assign the color. Now
after all of that is done we are ready to draw one side of the object. I
choose to use an API function to do this (Polygon (Window handler, 
coordinates of points drawn, number of sides of object)). I choose this 
method because all of the formulas I came up with that worked reliably
were really slow (even on a PII450). But how they work is basicaly like
so; find the top most point, find the next lowest point determine witch
side it's on, find slope of line connecting these two points, store all 
of the X values for this line in an array, find lowest point on opposite
side of object, determine side, store all of the X values for this line 
in an array, fill between these two lines by incrementing y by one and 
drawing a line from the corisponding x points in the arrays, and repeat 
untill object is filled. The Polygon function does all of this in one 
line. After drawing the side change fillstyle back to transparent so none
of the other objects get drawn in an odd way.

Well that is the quick explanation of how this program works. If you have 
any questions just follow the code I tried to comment it pretty well, but
if you still have questions e-mail me and I'll do my best to answer them.

            