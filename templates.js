/* SPARK models and evaluation templates — SINGLE SOURCE for assess.html and matrix.html.
 * backend/SPARK_Code.gs carries a generated copy (between the SPARK_TEMPLATES markers);
 * after editing this file, regenerate that block and redeploy the backend.
 * Positions matter: results are stored per model and per item in this order. Only ADD at the end of a list. */
const SPARK = {
 "criteria5": [
  "Assembly Quality & Accuracy",
  "Fastener Tightening & Torque Compliance",
  "Workmanship & Component Handling",
  "Safety & Standard Work Compliance",
  "Inspection & Non-Conformity Management"
 ],
 "templates": {
  "skd19": {
   "kind": "areas",
   "title": "SKD assembly – 19 key areas",
   "source": "SKD Assembly.xlsx",
   "areas": [
    {
     "name": "Depacking Process",
     "crit": [
      "Safety compliance",
      "Correct part identification",
      "Proper handling / damage prevention",
      "Work method & sequence",
      "Housekeeping"
     ]
    },
    {
     "name": "Trolly Loading Process",
     "crit": [
      "Correct loading sequence",
      "Load stability",
      "Safe handling",
      "Damage prevention",
      "Workplace housekeeping"
     ]
    },
    {
     "name": "Hoister Usage",
     "crit": [
      "Pre-use safety check",
      "Correct lifting point",
      "Safe operation",
      "Load balance / stability",
      "Post-use handling"
     ]
    },
    {
     "name": "Handle Bar Assembly",
     "crit": [
      "Correct part & orientation",
      "Proper fitment",
      "Torque / tightening",
      "Alignment",
      "Final quality check"
     ]
    },
    {
     "name": "Speedometer Assembly",
     "crit": [
      "Correct part identification",
      "Correct fitment",
      "Wiring / connector condition",
      "Function check",
      "Workmanship"
     ]
    },
    {
     "name": "HSG Head Lamp Assembly",
     "crit": [
      "Correct fitment",
      "Wiring / connector condition",
      "Alignment",
      "Fastener tightening",
      "Function check"
     ]
    },
    {
     "name": "TSL FR L & R Assembly",
     "crit": [
      "Correct left/right parts",
      "Correct fitment",
      "Wiring / connector condition",
      "Alignment",
      "Final inspection"
     ]
    },
    {
     "name": "Front Fender Assembly",
     "crit": [
      "Correct fitment",
      "Alignment",
      "Fastener tightening",
      "Damage prevention",
      "Final inspection"
     ]
    },
    {
     "name": "Wheel Assembly – Front",
     "crit": [
      "Correct wheel / component identification",
      "Axle & fitment",
      "Torque / tightening",
      "Alignment",
      "Brake / rotation check"
     ]
    },
    {
     "name": "Muffler Assembly",
     "crit": [
      "Correct fitment",
      "Fastener torque",
      "Clearance",
      "Leakage / abnormal noise check",
      "Final inspection"
     ]
    },
    {
     "name": "Crash Guard Assembly",
     "crit": [
      "Correct position",
      "Fitment",
      "Torque / tightening",
      "Clearance",
      "Final inspection"
     ]
    },
    {
     "name": "Main Footrest Assembly",
     "crit": [
      "Correct fitment",
      "Torque / tightening",
      "Movement / function",
      "Alignment",
      "Final inspection"
     ]
    },
    {
     "name": "Footrest L & R Assembly",
     "crit": [
      "Correct left/right parts",
      "Correct fitment",
      "Torque / tightening",
      "Movement / function",
      "Final inspection"
     ]
    },
    {
     "name": "Unloading",
     "crit": [
      "Safe handling",
      "Correct unloading sequence",
      "Damage prevention",
      "Material identification",
      "Housekeeping"
     ]
    },
    {
     "name": "Testing",
     "crit": [
      "Correct test procedure",
      "Safety compliance",
      "Defect identification",
      "Accurate result recording",
      "Final confirmation"
     ]
    },
    {
     "name": "Minor Repair",
     "crit": [
      "Correct diagnosis",
      "Correct repair method",
      "Proper tool usage",
      "Quality after repair",
      "Documentation / confirmation"
     ]
    },
    {
     "name": "Major Repair",
     "crit": [
      "Correct diagnosis",
      "Correct repair method",
      "Safety compliance",
      "Quality after repair",
      "Final verification"
     ]
    },
    {
     "name": "Identified Damage Parts & Shortages",
     "crit": [
      "Correct identification",
      "Accurate quantity",
      "Proper segregation",
      "Timely reporting",
      "Documentation"
     ]
    },
    {
     "name": "Quality of Work",
     "crit": [
      "First-time-right",
      "Workmanship",
      "Defect prevention",
      "Standard compliance",
      "Final quality awareness"
     ]
    }
   ]
  },
  "rtr_ws": {
   "kind": "stations",
   "title": "SKD workstation evaluation – RTR 160 4V / RTR 200",
   "source": "Test Revised with criteria.xlsx (RTR 160 , 200)",
   "stations": [
    {
     "code": "L-1",
     "items": [
      "Unpack the motorcycle parts from the box.",
      "Remove the polythene covers from the parts.",
      "Fit the fuel tank cover [safety cover] over the fuel tank before starting assembly.",
      "Inspect the parts for defects and report any findings to the Transit Team."
     ]
    },
    {
     "code": "L-2",
     "items": [
      "Place the front wheel, saree guard, crash guard, and seat inside the trolley.",
      "Remove the hardware box, separate the bolts, and arrange them in the tray.",
      "Send the left and right tail covers and front fender to the sub-assembly area."
     ]
    },
    {
     "code": "L-3",
     "items": [
      "Remove the wooden piece fitted at the handle mounting area.",
      "Handle assembly.",
      "Install the left pillion footrest.",
      "Assemble the front brake master pump.",
      "Install the rear number plate.",
      "Remove the crash guard and center-stand mounting brackets from the main wooden plate.",
      "Use the hoist to lift the motorcycle and place it on the two-wheeler assembly trolley.",
      "Insert the rear shock absorber bolt (14 mm).",
      "Remove the rear-wheel mounting bracket and the front-wheel axle.",
      "Fix the center stand.",
      "Secure the rear wheel to the two-wheeler assembly trolley using the ratchet."
     ]
    },
    {
     "code": "SUB Process - 1",
     "items": [
      "Fit the left and right tail covers to the center tail cover.",
      "Fit the tail lamp to the complete tail-cover assembly."
     ]
    },
    {
     "code": "SUB Process - 2",
     "items": [
      "Carry out the engine guard sub-assembly.",
      "Carry out the front fender sub-assembly."
     ]
    },
    {
     "code": "L-4",
     "items": [
      "Fit and adjust the throttle cable.",
      "Pre-fit the crash guard to the motorcycle.",
      "Fit the tail-cover assembly to the motorcycle.",
      "Open the right frame cover and connect the battery terminals.",
      "Tighten the center-stand fasteners.",
      "Tighten the rear shock absorber fasteners correctly.",
      "Install the prop stand and ensure the stand switch is fitted correctly.",
      "Fit and tighten the saree guard correctly."
     ]
    },
    {
     "code": "L-5",
     "items": [
      "Tighten the crash guard fasteners correctly.",
      "Tighten the front brake master pump fasteners correctly.",
      "Tighten the frame-cover fasteners correctly.",
      "Install the speedometer correctly.",
      "Fit the headlamp housing (HSG) correctly.",
      "Install the grab handle correctly."
     ]
    },
    {
     "code": "L-6",
     "items": [
      "Install the muffler assembly correctly.",
      "Install the front wheel correctly and mark the fastener after torque tightening.",
      "Install the front fender correctly.",
      "Fit the engine guard securely to the motorcycle.",
      "Install the front brake caliper correctly onto the front wheel disc assembly.",
      "Route and connect the ABS system wiring correctly.",
      "Fit the speedometer cover bracket correctly to the motorcycle.",
      "Install the front number plate and tighten the rear number plate securely.",
      "Remove the fuel tank cover and release the ratchet.",
      "Unload the motorcycle for testing."
     ]
    },
    {
     "code": "Testing",
     "items": [
      "Carry out the vehicle tests."
     ]
    },
    {
     "code": "Repair",
     "items": [
      "Carry out major and minor repairs as required."
     ]
    }
   ]
  },
  "ckd_xl100": {
   "kind": "stations",
   "title": "CKD workstation evaluation – XL 100",
   "source": "Test Revised with criteria.xlsx (CKD)",
   "stations": [
    {
     "code": "L-1",
     "items": [
      "Attach stay head lamp bracket to the fork legs",
      "Fix fork assy to the frame",
      "Fix handle lock to the frame",
      "Attach sub assembled handle bar and Tighten the steering lock nut",
      "Attach and tighten cap nut to handle column",
      "Arrange wiring harness through the H bracket and attach cable guide LH",
      "Load vehicle to the conveyor",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "L-2",
     "items": [
      "Place Chassis on the work table and Apply grease on the racers and place cup set balls",
      "Attach Front seat on the frame",
      "Route wiring harness on the chassis",
      "Fix horn with front number plate bracket to the vehicle",
      "Connect meter wire sockets with the main wiring harness",
      "Attach rear fender to the vehicle and attach seat mounting nuts",
      "Attach rear fender front side bottom bolt",
      "Attach air cleaner assembly to the vehicle and attach cable guide RH to the frame",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "L-3",
     "items": [
      "Attach shock absorber to frame",
      "Fix tail lamp to vehicle and Reflector and mud flap assembly",
      "Attach swing arm to the frame",
      "Attach Rear drum sprocket to the swing arm",
      "Tighten both side attached front seat mounting nuts",
      "Cover tool box center rear assembly",
      "Attach luggage carrier mounting bolts to the frame",
      "Attach Pillion Handle to the luggage",
      "Cushion rubber attach to the bracket of the frame",
      "Attach rear wheel to the swing arm assist with other side operator",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "L-4",
     "items": [
      "Attach rear fender to the frame bracket",
      "Attach luggage carrier RH mounting screws",
      "Fix Rear wheel to the vehicle",
      "Attach torque link with the rear brake panel",
      "Shock absorber RH and Lady handle assembly; attach Pillion Handle to the luggage",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "L-5",
     "items": [
      "Attach engine mount bolts",
      "Tighten cover tool box center and luggage mounting screws",
      "Battery assembly to the vehicle",
      "Side stand place to the frame and attach bolt",
      "Tighten side stand fixing bolt",
      "Fix SAI unit to the engine",
      "Front chain sprocket assembly",
      "Route rear brake cable through head lamp bracket",
      "Fix tail lamp to signal lamp bracket and Attach tail lamp to the vehicle",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "L-6",
     "items": [
      "Fix engine to the vehicle and fix Ignition coil mounting bracket to the vehicle",
      "Route engine wire through the tool box bottom and attach cover tool box center rear RH screws",
      "Tighten engine mount bolts and tighten kick lever mounting bolt",
      "Control unit attach to the vehicle and connect wire clips",
      "Mount ignition coil to the chassis, attach engine ground wire",
      "Route choke cable and Acc. Cable through the chassis bracket and air cleaner breather hose attach to the engine",
      "Attach RH side tail lamp bracket bolt",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "L-7",
     "items": [
      "Take the kitting bin and muffler from the trolley and it place on the conveyor",
      "Tighten the drum sprocket nut",
      "Fix pillion footrest LH and Tighten shock absorber top bolt and bottom nut",
      "Attach screw clips and beading to cover tool box and place it on tool box rear",
      "Fix tool box RH screws and Attach and tighten air cleaner mount screw",
      "Attach front fender mounting RH side screws",
      "Attach front brake cable with brake panel and attach choke cable to lever",
      "Fix sari guard to the vehicle",
      "Chain case assembly",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "L-8",
     "items": [
      "Attach fender front RH mounting screws",
      "Attach split pin and Tighten rear wheel axle nut",
      "Fix muffler to the vehicle",
      "Removed head lamp place on the conveyor",
      "Tighten tool box mounting RH screws and Tighten air cleaner bellow screw",
      "Attach rear brake cable and adjust rear brake as required",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "L-9",
     "items": [
      "Attach fuel tank on the frame",
      "Tighten tail lamp mounting bracket screws and fender front mounting screws",
      "Fix engine sprocket cover",
      "Tighten fuel tank mounting screws",
      "Attach cover frame to frame",
      "Attach head lamp housing to the stay head lamp bracket",
      "Connect all wire sockets and connectors in head lamp housing",
      "Fix seat lock to the carrier",
      "Fix head lamp to the head lamp rear",
      "Cable clip attach with the rear brake cable and choke cable",
      "Adjust the front brake",
      "Take the empty hardware bin on the conveyor and it place on the side trolley",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "L-10",
     "items": [
      "Attach acc. Cable top end to the acc. Grip and tighten control switch bolts",
      "Fix pilion foot rest RH and Tighten shock absorber bolts and Tighten swing arm nut",
      "Attach and tighten both side floor board mounting screws and Tighten fender front stabiliser tube RH screws",
      "Fix leg guard to the vehicle",
      "Connect LH/RH TSL couplers and attach cable strip for handle switch wire",
      "Fix tool box cover lock to the utility box",
      "Attach and tighten head lamp mounting screw and fix rear seat",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "Testing",
     "items": [
      "Place vehicle on the testing bay",
      "Follow the tester's check list for test the vehicle",
      "Rectify defects which can be rectify within the stage time",
      "Remove fuel hose from tank and connect testing fuel can tube and fill fuel to the carb sump",
      "Start the vehicle using self starter and do the necessary checks",
      "Start the vehicle using kick starter",
      "Do the test run",
      "If found any defect which cant be rectified within stage time, update the defect card",
      "Park vehicle on the relevant work station & return back to the testing area"
     ]
    }
   ]
  },
  "w3_ws": {
   "kind": "stations",
   "title": "3W workstation evaluation – 3W / EV Max 3W",
   "source": "Test Revised with criteria.xlsx (3Wheel)",
   "stations": [
    {
     "code": "W-1",
     "items": [
      "3W structure tilting and Depacking",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "W-2",
     "items": [
      "Prepare for assembling",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "W-3",
     "items": [
      "Remove battery from the vehicle & send to charging operation",
      "Fit charged battery to the vehicle",
      "Fix side boards, passenger seat & passenger backrest to the vehicle",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "W-4",
     "items": [
      "Remove driver seat mounting bracket",
      "Fix driver seat mounting bracket",
      "Check tightness of the passenger compartment mounting bolts",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "W-5",
     "items": [
      "Pre fit tail lamps LHS/RHS & number plate lamp",
      "Fix rear bumper to the vehicle",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "W-6",
     "items": [
      "Fix rear shock absorber and fix Inner Mud Flap LHS",
      "Pre-fit rear inner mudguard assy",
      "Fix rear LH side mud flap to the vehicle",
      "Fix rear hood bar cage & driver backrest to the vehicle",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "W-7",
     "items": [
      "Fix rear shock absorber and fix Inner Mud Flap RHS",
      "Pre-fit rear inner mudguard assy",
      "Fix rear RH side mud flap to the vehicle and rear number plate to the vehicle",
      "Fix rear hood bar cage & driver backrest to the vehicle",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "W-8",
     "items": [
      "Fit steering column & front wheel to the vehicle and Tighten steering column & fit handle bar",
      "Tighten the front wheel mounting nut and the handle mounting bolt",
      "Fix footrest to flow board",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "W-9",
     "items": [
      "Pre fit wind shield frame",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "W-10",
     "items": [
      "Fix wiper motor, wiper blade & front number plate to the vehicle",
      "Fix windshield frame brackets to the vehicle",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "W-11",
     "items": [
      "Fix windshield apron plastic cover",
      "Attach plastic cover to utility box",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "W-12",
     "items": [
      "Attach the center hood bar to the vehicle",
      "Fix side hood bar LHS to the vehicle",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "Testing",
     "items": [
      "Remove air bubbles in brake system and Unload the vehicle from 3W dolley",
      "Fill oil to the engine & differential",
      "Fuel filling & wheel nut tightening",
      "Vehicle testing"
     ]
    },
    {
     "code": "Hood",
     "items": [
      "Fix hood strips to the vehicle",
      "Fix hood timber to the vehicle",
      "Hood making-A",
      "Hood making-B",
      "Hood making-C",
      "Remove the LH side door lock and side door lock bracket (DVA)",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "Repair",
     "items": [
      "Minor and Major rectification work"
     ]
    }
   ]
  }
 },
 "lines": [
  {
   "id": "SKD",
   "name": "2W SKD",
   "models": [
    "Sport 110",
    "Ntorq 125",
    "NTORQ 150",
    "Ronin 225",
    "RTR 160 4V",
    "RTR 200",
    "RTR 310",
    "RTR 310 RR",
    "iQube",
    "Orbiter",
    "Jupiter",
    "Raider 125"
   ]
  },
  {
   "id": "CKD",
   "name": "2W CKD",
   "models": [
    "XL 100 i-Touch"
   ]
  },
  {
   "id": "3W",
   "name": "3W",
   "models": [
    "3W",
    "EV Max 3W"
   ]
  }
 ],
 "modelTemplate": {
  "Sport 110": "skd19",
  "Ntorq 125": "skd19",
  "NTORQ 150": "skd19",
  "Ronin 225": "skd19",
  "RTR 160 4V": "rtr_ws",
  "RTR 200": "rtr_ws",
  "RTR 310": "skd19",
  "RTR 310 RR": "skd19",
  "iQube": "skd19",
  "Orbiter": "skd19",
  "Jupiter": "skd19",
  "Raider 125": "skd19",
  "XL 100 i-Touch": "ckd_xl100",
  "3W": "w3_ws",
  "EV Max 3W": "w3_ws"
 }
};
