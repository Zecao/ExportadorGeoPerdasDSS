

namespace AuxClasses
{
    class PVSystemPar
    {
        public bool _modelPVSystems; // false = exports PVSystems as generator model=1
        public bool _geraInvControl;
        public string _invControlMode;
        public string _voltVarcurve;
        public string _varFollowInv; // set False to night mode (and if VOLTVAR)

        /*
        _modelPVSystems     -> False = exports PVSystems as generator model = 1
        _invControlModeMV   -> "PF=1" "VOLTVAR" "voltwatt"
        _varFollowInvMV     -> False to night mode(and if VOLTVAR)
        _voltVarcurve =     -> "voltvar_c";  ///uses voltvar_0 to no voltvar
        */

        public PVSystemPar(bool modelPVSystemsLV, bool gerIC, string invControlModeLV = "VOLTVAR", string varFollowInvLV = "False", string vvarCurve = "voltvar_c")
        {
            _modelPVSystems = modelPVSystemsLV;
            _geraInvControl = gerIC;

            // Optional parameters 
            _invControlMode = invControlModeLV;
            _varFollowInv = varFollowInvLV;
            _voltVarcurve = vvarCurve;
        }

        public bool GeraInvControl()
        {
            return _geraInvControl;
            /*
            if ( _modelPVSystems) //&& ! _invControlMode.Equals("")
                return true;
            else
                return false;
                */
        }
    }
}
