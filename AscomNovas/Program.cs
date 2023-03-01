// See https://aka.ms/new-console-template for more information

using Microsoft.VisualBasic;
using System.ComponentModel;
using System.Dynamic;
using System.Reflection;
using System.Security.Cryptography;
using System;
using System.Collections.Generic;
using System.IO;



Console.WriteLine("Hello, World!");
/*
Planet planet = new Planet();
Ephemeris ephemeris = new Ephemeris() ;
string Elements = "";
    planet.Ephemeris = ephemeris;
  
    if (Elements.Length < 50)
    {

    }
    else
    {
        try
        {
            planet.Name = Elements.Substring(0, 7).Trim();
            ephemeris.Epoch = SkyMath._packToJulian(Elements.Substring(20, 5).Trim());
            ephemeris.M = double.Parse(Elements.Substring(26, 9).Trim());
            ephemeris.n = double.Parse(Elements.Substring(80, 11).Trim());
            ephemeris.a = double.Parse(Elements.Substring(92, 11).Trim());
            ephemeris.e = double.Parse(Elements.Substring(70, 9).Trim());
            ephemeris.Peri = double.Parse(Elements.Substring(37, 9).Trim());
            ephemeris.Node = double.Parse(Elements.Substring(48, 9).Trim());
            ephemeris.Incl = double.Parse(Elements.Substring(59, 9).Trim());
        }
        catch (Exception ex)
        {
            Name = ex.Message;
            RA = Dec = 0.0;
            Marshal.ReleaseComObject((object)planet);
            Marshal.ReleaseComObject((object)ephemeris);
          
        }
    }
    try
    {
        planet.Number = int.Parse(planet.Name);
    }
    catch (FormatException ex)
    {
        planet.Number = 100000;
    }
    PositionVector astrometricPosition = planet.GetAstrometricPosition(SkyMath.DateUtcToJulian(tUtc));
    RA = astrometricPosition.RightAscension;
    Dec = astrometricPosition.Declination;



*/



dynamic Utl = COMObject.CreateObject("ASCOM.Utilities.Util");
dynamic Novas = COMObject.CreateObject("ASCOM.Astrometry.NOVAS.NOVAS31");
//dynamic Kepler = COMObject.CreateObject("ASCOM.Kepler");
/*
 * 
 * 
 * pl = CreateObject("ASCOM.Astrometry.NOVASCOM.Planet")
                    kt = CreateObject("ASCOM.Astrometry.Kepler.Ephemeris")
                    ke = CreateObject("ASCOM.Astrometry.Kepler.Ephemeris")
                    Earth = CreateObject("ASCOM.Astrometry.NOVASCOM.Earth")
                    Ut = CreateObject("ASCOM.Utilities.Util")
                    Site = CreateObject("ASCOM.Astrometry.NOVASCOM.Site")
                    'AstroUtils = CreateObject("ASCOM.Astrometry.AstroUtils.AstroUtils")
                    AstroUtils = CreateObject("ASCOM.Astrometry.NOVAS.NOVAS31")*/

//var Utl = new ASCOM.Utilities.Util(); // Get current Julian date
var JD = Utl.JulianDate;

dynamic NCStar = COMObject.CreateObject("ASCOM.Astrometry.NOVASCOM.Star"); // new Star(); // Create NOVASCOM objects
dynamic NCEarth = COMObject.CreateObject("ASCOM.Astrometry.NOVASCOM.Earth");
dynamic NCPlanet = COMObject.CreateObject("ASCOM.Astrometry.NOVASCOM.Planet");

dynamic NCAsteroid = COMObject.CreateObject("ASCOM.Astrometry.NOVASCOM.Planet");
NCAsteroid.Type = ASCOM.Astrometry.BodyType.MinorPlanet;

dynamic ephem = COMObject.CreateObject("ASCOM.Astrometry.Kepler.Ephemeris");
ephem.BodyType = ASCOM.Astrometry.BodyType.MinorPlanet;



dynamic NCSite = COMObject.CreateObject("ASCOM.Astrometry.NOVASCOM.Site");
NCSite.Set(0, 0, 1500); // Lat, lon, altitude
Console.WriteLine("Lat is " + NCSite.Latitude);
/*NCSite.Height = 0; // Initialise site object
NCSite.Latitude = 0;
NCSite.Longitude = 0;
NCSite.Pressure = 1000;
NCSite.Temperature = 10; // .0d)*/
dynamic NCPositionVector = COMObject.CreateObject("ASCOM.Astrometry.NOVASCOM.PositionVector");

NCPositionVector.SetFromSite(NCSite, (double)11); // Find site position cartesian co-ordinates, throws exception
Console.WriteLine("NOVAS.COM SetFromSite " + NCPositionVector.x + " " + NCPositionVector.y + " " + NCPositionVector.z + " " + NCPositionVector.LightTime);

NCStar.Set(9.0d, 25.0d, 0.0d, 0.0d, 0.0d, 0.0d); // Initialise the star object
NCPositionVector = NCStar.GetAstrometricPosition(JD); // Find astrometric and topocentric right ascension and Declination
var RA = NCPositionVector.RightAscension; //Utl.HoursToHMS(NCPositionVector.RightAscension);
Console.WriteLine("NOVAS.COM Star Astrometric Position " + RA + " " + Utl.DegreesToDMS(NCPositionVector.Declination, ":", ":"));
NCPositionVector = NCStar.GetTopocentricPosition(JD, NCSite, false);
Console.WriteLine("NOVAS.COM Star Topocentric Position " + Utl.HoursToHMS(NCPositionVector.RightAscension) + " " + Utl.DegreesToDMS(NCPositionVector.Declination, ":", ":"));

NCEarth.SetForTime(JD); // Initialise earth object
NCPositionVector = NCEarth.BarycentricPosition; // Find its barycentric and heliocentric position
Console.WriteLine("NOVAS.COM Earth BaryPos x ", NCEarth.BarycentricPosition.x);
Console.WriteLine("NOVAS.COM Earth BaryPos y ", NCEarth.BarycentricPosition.y);
Console.WriteLine("NOVAS.COM Earth BaryPos z ", NCEarth.BarycentricPosition.z);
Console.WriteLine("NOVAS.COM Earth HeliPos x ", NCEarth.HeliocentricPosition.x);
Console.WriteLine("NOVAS.COM Earth HeliPos y ", NCEarth.HeliocentricPosition.y);
Console.WriteLine("NOVAS.COM Earth HeliPos z ", NCEarth.HeliocentricPosition.z);

Console.WriteLine("NOVAS.COM Barycentric Time ", NCEarth.BarycentricTime); // Find other ephemeris information
Console.WriteLine("NOVAS.COM Equation Of Equinoxes ", NCEarth.EquationOfEquinoxes);
Console.WriteLine("NOVAS.COM Mean Obliquity ", NCEarth.MeanObliquity);
Console.WriteLine("NOVAS.COM Nutation in Longitude ", NCEarth.NutationInLongitude);
Console.WriteLine("NOVAS.COM Nutation in Obliquity ", NCEarth.NutationInObliquity);
Console.WriteLine("NOVAS.COM True Obliquity ", NCEarth.TrueObliquity);

NCPlanet.Name = "Saturn"; // Initialise Planet object as Saturn
NCPlanet.Number = 6;
NCPlanet.Type = ASCOM.Astrometry.BodyType.MajorPlanet;
NCPositionVector = NCPlanet.GetAstrometricPosition(JD); // Find Saturn's astrometric position
Console.WriteLine("NOVAS.COM Saturn Astrometric Poistion " + NCPositionVector.x + " " + NCPositionVector.y + " " + NCPositionVector.z + " " + NCPositionVector.LightTime);
NCPositionVector = NCPlanet.GetTopocentricPosition(JD, NCSite, false); // Find Saturn's topocentric position
Console.WriteLine("NOVAS.COM Saturn Topocentric Poistion " + NCPositionVector.x + " " + NCPositionVector.y + " " + NCPositionVector.z + " " + NCPositionVector.LightTime);
Console.ReadLine();

class COMObject : DynamicObject
{
    private readonly object instance;

    public static COMObject CreateObject(string progID)
    {
        return new COMObject(Activator.CreateInstance(Type.GetTypeFromProgID(progID, true)));
    }

    public COMObject(object instance)
    {
        this.instance = instance;
    }

    public override bool TryGetMember(GetMemberBinder binder, out object result)
    {
        result = instance.GetType().InvokeMember(
            binder.Name,
            BindingFlags.GetProperty,
            Type.DefaultBinder,
            instance,
            new object[] { }
        );
        return true;
    }

    public override bool TrySetMember(SetMemberBinder binder, object value)
    {
        instance.GetType().InvokeMember(
            binder.Name,
            BindingFlags.SetProperty,
            Type.DefaultBinder,
            instance,
            new object[] { value }
        );
        return true;
    }

    public override bool TryInvokeMember(InvokeMemberBinder binder, object[] args, out object result)
    {
        result = instance.GetType().InvokeMember(
            binder.Name,
            BindingFlags.InvokeMethod,
            Type.DefaultBinder,
            instance,
            args
        );
        return true;
    }
}
