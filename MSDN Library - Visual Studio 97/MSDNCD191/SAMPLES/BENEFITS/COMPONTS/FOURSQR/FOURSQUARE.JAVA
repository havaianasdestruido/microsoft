//******************************************************************************
// foursquare.java:	Applet
//
//******************************************************************************
import java.applet.*;
import java.awt.*;

//==============================================================================
// Main Class for applet foursquare
//
//==============================================================================
public class foursquare extends Applet
{

	// PARAMETER SUPPORT:
	//		Parameters allow an HTML author to pass information to the applet;
	// the HTML author specifies them using the <PARAM> tag within the <APPLET>
	// tag.  The following variables are used to store the values of the
	// parameters.
    //--------------------------------------------------------------------------

    // Members (some for applet parameters)
    // <type>       <MemberVar>    = <Default Value>
    //--------------------------------------------------------------------------
	private float m_maxben;
	private int m_bgcolor = 0x00666666;
	private int m_fgcolor = 0x00FFFFFF;
	private float m_pre = m_maxben;
	private float m_post = 0;
	private String m_prelabel;
	private String m_postlabel;
	private String m_currency;
	private String m_midlabel;
	private Font labelFont;
	private Font bigFont;
	private Color m_darkRed;

    // Parameter names.  To change a name of a parameter, you need only make
	// a single change.  Simply modify the value of the parameter string below.
    //--------------------------------------------------------------------------
	private final String PARAM_maxben = "maxben";
	private final String PARAM_bgcolor = "bgcolor";
	private final String PARAM_fgcolor = "fgcolor";
	private final String PARAM_prelabel = "prelabel";
	private final String PARAM_postlabel = "postlabel";
	private final String PARAM_currency = "currency";
	private final String PARAM_midlabel = "midlabel";
	// foursquare Class Constructor
	//--------------------------------------------------------------------------
	public foursquare()
	{
		// TODO: Add constructor code here
	}

	// APPLET INFO SUPPORT:
	//		The getAppletInfo() method returns a string describing the applet's
	// author, copyright date, or miscellaneous information.
    //--------------------------------------------------------------------------
	public String getAppletInfo()
	{
		return "Name: foursquare\r\n" +
		       "Author: C.P.H. Lewis\r\n" +
		       "Created with Microsoft Visual J++ Version 1.1";
	}

	// PARAMETER SUPPORT
	//		The getParameterInfo() method returns an array of strings describing
	// the parameters understood by this applet.
	//
    // foursquare Parameter Information:
    //  { "Name", "Type", "Description" },
    //--------------------------------------------------------------------------
	public String[][] getParameterInfo()
	{
		String[][] info =
		{
			{ PARAM_maxben, "float", "Annual benefit total per employee" },
			{ PARAM_bgcolor, "int", "Background color; use decimal value of hex RRGGBB" },
			{ PARAM_fgcolor, "int", "Foreground color; use decimal value of hex RRGGBB" },
			{ PARAM_prelabel, "String", "Label for left half of graph" },
			{ PARAM_postlabel, "String", "Label for right half of graph" },
			{ PARAM_currency, "String", "Currency character" },
			{ PARAM_midlabel, "String", "Label for middle of graph" }
		};
		return info;		
	}

	// The init() method is called by the AWT when an applet is first loaded or
	// reloaded.  Override this method to perform whatever initialization your
	// applet needs, such as initializing data structures, loading images or
	// fonts, creating frame windows, setting the layout manager, or adding UI
	// components.
    //--------------------------------------------------------------------------
	public void init()
	{
		// PARAMETER SUPPORT
		//		The following code retrieves the value of each parameter
		// specified with the <PARAM> tag and stores it in a member
		// variable.
		//----------------------------------------------------------------------
		String param;

		// maxben: Annual benefit total per employee
		// bgcolor: Sets background color 
		// fgcolor: Sets color of labels and axes
		// prelabel: Describes left-hand side
		// postlabel: Describes right-hand side
		// midlabel: Describes the x-axis
		// currency: the character for the local currency
		//----------------------------------------------------------------------
		param = getParameter(PARAM_maxben);
		if (param != null)
			m_maxben = Float.valueOf(param).floatValue();
		param = getParameter(PARAM_bgcolor);
		if (param != null)
			m_bgcolor = Integer.valueOf(param).intValue();
		param = getParameter(PARAM_fgcolor);
		if (param != null)
			m_fgcolor = Integer.valueOf(param).intValue();
		param = getParameter(PARAM_prelabel);
		if (param != null)
			m_prelabel = param;
		else
			m_prelabel = "";
		param = getParameter(PARAM_postlabel);
		if (param != null)
			m_postlabel = param;
		else
			m_postlabel = "";
		param = getParameter(PARAM_currency);
		if (param != null)
			m_currency = param;
		else
			m_currency = "";
		param = getParameter(PARAM_midlabel);
		if (param != null)
			m_midlabel = param;
		else
			m_midlabel = "";

        // If you use a ResourceWizard-generated "control creator" class to
        // arrange controls in your applet, you may want to call its
        // CreateControls() method from within this method. Remove the following
        // call to resize() before adding the call to CreateControls();
        // CreateControls() does its own resizing.
        //----------------------------------------------------------------------
		Color bg, fg;
		bg = new Color(m_bgcolor);
		fg = new Color(m_fgcolor);
		setBackground(bg);
		setForeground(fg);
		labelFont = new Font("Helvetica", Font.BOLD, 14);
		bigFont = new Font("Helvetica", Font.BOLD, 36);
		m_darkRed = new Color(0x00880000);
		m_pre = m_maxben;
		repaint(); // to set the background correctly

	}

	// Place additional applet clean up code here.  destroy() is called when
	// when you applet is terminating and being unloaded.
	//-------------------------------------------------------------------------
	public void destroy()
	{
		// TODO: Place applet cleanup code here
	}


	// foursquare Paint Handler
	//--------------------------------------------------------------------------
	public void paint(Graphics g)
	{
		Rectangle r = bounds();
		int x1, y1;

		// we need room for two lines of text at the top, then the rest is 
		// divided evenly 



		g.translate( r.width/2, (r.height/2) + 16); // setting the middle of the graphic part to 0,0

		y1 = (r.height - 32) / 2;
		x1 = r.width / 3;

		// sketch in the outline of the max benefit credit

		g.setColor(Color.gray);
		g.drawLine( -x1 , 0, -x1 , -y1);
		g.drawLine( -x1 , -y1, 0, -y1);

		g.drawLine(- (3 * x1 / 2), 0, (3 * x1 / 2), 0);
		g.drawLine(0, -y1, 0, y1);

		// we draw the rectangles proportionately to the total benefit - but 
		// clip the rectangles if they threaten to overlap the labels
		// ( at the bottom edge, they'll vanish over the edge of the drawing region)

		int prePortion, postPortion;
		if ( m_pre >= m_maxben )
			prePortion = y1;
		else
			prePortion =  (int)( m_pre  / m_maxben * y1);
		if (m_post >= m_maxben )
			postPortion = y1;
		else
			postPortion = (int)( m_post / m_maxben * y1);

		// draw the rectangles - 3D highlights exaggerated 

		g.setColor(Color.darkGray);

		if ( prePortion > 0)
			g.fillRect(-x1, -prePortion, x1, prePortion);
		if (postPortion > 0)
			g.fillRect( 1, -postPortion, x1, postPortion);

		g.setColor(Color.black);
		if ( prePortion > 0)
			g.fill3DRect(-x1 + 2, -prePortion + 2, x1 - 2, prePortion - 2, true);
		if (postPortion > 0)
			g.fill3DRect( 3, -postPortion + 2, x1 -2, postPortion - 2, true);

		g.setColor(Color.red);

		if ( prePortion < 0)
			g.fillRect( -x1, 0, x1, -prePortion);
		if ( postPortion < 0)
			g.fillRect( 1, 0, x1, -postPortion);

		g.setColor(m_darkRed);

		if ( prePortion < 0)
			g.fill3DRect( -x1 + 1, 1, x1 -1, -prePortion -1, true);
		if ( postPortion < 0)
			g.fill3DRect( 2, 1, x1 -1, -postPortion -1, true);

		g.setColor(getForeground());

		// Labels, for persons more textually inclined.
		g.setFont(labelFont);
		g.drawString(m_prelabel, -( 3 * r.width / 8), -y1 - 16);
		if ( m_pre >= 0 )
			g.drawString(m_currency + Float.toString(m_pre), -(3 *  r.width / 8), -y1 - 2);
		else
			g.drawString("( " + m_currency + Float.toString(-m_pre) + ")", -(3 *  r.width / 8), -y1 - 2);
		g.drawString(m_postlabel, ( r.width / 8),  - y1 - 16);
		if ( m_post >= 0)
			g.drawString(m_currency + Float.toString(m_post), ( r.width / 8), -y1 - 2);
		else
			g.drawString("( " + m_currency + Float.toString(-m_post) + " )", ( r.width / 8), -y1 - 2);
		// should center this by finding out the length of the string - 
		// still have to guess at font proportions...
		g.drawString(m_midlabel, -( m_midlabel.length() * 5), 5);
		g.setFont(bigFont);
		g.drawString("+", -8, -y1 + 36);
		g.drawString("-", -7, y1 - 16);

	}

	//		The start() method is called when the page containing the applet
	// first appears on the screen. The AppletWizard's initial implementation
	// of this method starts execution of the applet's thread.
	//--------------------------------------------------------------------------
	public void start()
	{
		// TODO: Place additional applet start code here
	}
	
	//		The stop() method is called when the page containing the applet is
	// no longer on the screen. The AppletWizard's initial implementation of
	// this method stops execution of the applet's thread.
	//--------------------------------------------------------------------------
	public void stop()
	{
	}


	// MOUSE SUPPORT:
	//		The mouseDown() method is called if the mouse button is pressed
	// while the mouse cursor is over the applet's portion of the screen.
	//--------------------------------------------------------------------------
	public boolean mouseDown(Event evt, int x, int y)
	{
		// give a label explaing the effect of this quadrant on your checkbook
		return true;
	}

	// MOUSE SUPPORT:
	//		The mouseUp() method is called if the mouse button is released
	// while the mouse cursor is over the applet's portion of the screen.
	//--------------------------------------------------------------------------
	public boolean mouseUp(Event evt, int x, int y)
	{
		// TODO: Place applet mouseUp code here
		return true;
	}

	public void recalc(float pretax, float posttax)
	{
		// all this needs is two signed numbers: 
		// the change to the pre-tax paycheck adjustment, 
		// and the change to the post-tax paycheck adjustment.
		m_pre += pretax;
		m_post += posttax;
		repaint();

	}
	// TODO: Place additional applet code here
	public void reset()
	{
	   // redraw with default data;
	   // useful for going forwards and backwards in a range of pages
	   m_pre = m_maxben;
	   m_post = 0;
	   repaint();
	}
}
