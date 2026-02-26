using System.Collections.ObjectModel;

namespace arv_mu_computation.My;

public class grfData
{
	public string name { get; set; }

	public Collection<datevalue> pssOn { get; set; }

	public Collection<datevalue> pssOff { get; set; }

	public double sPoint { get; set; }
}
