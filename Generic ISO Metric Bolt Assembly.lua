-- Revit Families ISO Metric Bolt & Bolt Assembly ------------------------------
-- Type catalog and lookup table generator

--- Initializes the global variables
function init()

  -- Paths and file names
  path = "C:/Users/wolff/btu/Cloud/LZKI-Gebäude/Bibliothek/Familien/Beschläge/Metrische Schrauben"
  fna  = "Generic ISO Metric Bolt Assembly.txt"
  fnb  = "Generic ISO Metric Bolt.txt"
  fngl = "ISO Metric Bolt Assembly -- Grip to Length.csv"
  
  -- Revit parameter types
  tlmm = "##LENGTH##MILLIMETERS"
  toth = "##OTHER##"

  -- Parameters in Revit type catalog of bolt assemblies
  tcanN = "[Name]"
  tcanD = "Nominal Diameter"; tcatD = tlmm
  tcanL = "Grip Length";      tcatL = tlmm
  tcanS = "Shank";            tcatS = toth
  tcanW = "Large Washers";    tcatW = toth
  tcanM = "Material";         tcatM = toth
  tcanT = "Thread Material";  tcatT = toth

  -- Parameters in Revit type catalog of bolts
  tcbnN = "[Name]"
  tcbnD = "Nominal Diameter"; tcbtD = tlmm
  tcbnL = "Length";           tcbtL = tlmm
  tcbnS = "Shank";            tcbtS = toth
  tcbnW = "Large Washer";     tcbtW = toth
  tcbnM = "Material";         tcbtM = toth
  tcbnT = "Thread Material";  tcbtT = toth

  -- Parameters in grip-to-length lookup table
  --     Parameter Name             Parameter Type
  lgnD  = "D";                lgtD  = tlmm
  lgnG  = "LG";               lgtG  = tlmm
  lgnL  = "l";                lgtL  = tlmm
  
  -- Materials
  materials = {
  --  "Steel", 
  "Steel galvanized", 
  --  "Steel galvanized yellow-chromated",
  --  "Stainless steel"
  }
  
  -- Nominal diameters (M<D>) 
  diameters = {3,4,5,6,8,10,12,14,16,18,20,22,24,27,30,33,36,39,42,45,48,52,56,64}
  
  -- Standard bolt lengths
  lengths = {}
  lengths[ 3] = {3,4,5,6,8,10,12,16,18,20,22,25,30,35,40,50,60}
  lengths[ 4] = {4,6,8,10,12,14,16,18,20,22,25,30,35,40,45,50,55,60,65,70,75,80}
  lengths[ 5] = {6,8,10,12,14,16,18,20,22,25,30,35,40,45,50,55,60,65,70,80,90,100}
  lengths[ 6] = {6,8,10,12,14,16,18,20,22,25,28,30,35,40,45,50,55,60,65,70,75,80,85,90,100,110,120,130,140,150}
  lengths[ 8] = {8,10,12,14,16,18,20,22,25,30,35,40,45,50,55,60,65,70,75,80,85,90,95,100,110,120,130,140,150,160,170,180,190,200}
  lengths[10] = {10,12,16,18,20,22,25,28,30,35,40,45,50,55,60,65,70,75,80,85,90,100,110,120,130,140,150,160,170,180,190,200,220,240,280,300}
  lengths[12] = {10,12,16,18,20,22,25,28,30,35,40,45,50,55,60,65,70,75,80,85,90,100,110,120,130,140,150,160,170,180,190,200,220,240,300}
  lengths[14] = {16,20,25,30,35,40,45,50,55,60,65,70,75,80,90,100,110,120,130,140,150,160,170,180,200,220}
  lengths[16] = {12,16,20,25,30,35,40,45,50,55,60,65,70,75,80,85,90,95,100,110,120,130,140,150,160,170,180,190,200,210,220,230,240,250,260,280,300,320,340,400,500}
  lengths[18] = {20,25,30,35,40,45,50,55,60,65,70,75,80,85,90,100,110,120,130,140,150,160,170,180,190,200}
  lengths[20] = {20,25,30,35,40,45,50,55,60,65,70,75,80,85,90,100,110,120,130,140,150,160,170,180,190,200,210,220,230,240,250,260,280,300,360}
  lengths[22] = {30,35,40,45,50,55,60,65,70,75,80,90,100,110,120,130,140,150,160,170,180,190,200}
  lengths[24] = {25,30,35,40,45,50,55,60,65,70,75,80,85,90,100,110,120,130,140,150,160,170,180,190,200,210,220,230,240,250,260,280,300,320,500}
  lengths[27] = {30,40,45,50,55,60,65,70,75,80,85,90,100,110,120,130,140,150,160,170,180,190,200,300}
  lengths[30] = {35,40,45,50,55,60,65,70,75,80,85,90,100,110,120,130,140,150,160,170,180,190,200,210,220,230,240,250,260,280,300,320,340,360,380,400,500,600}
  lengths[33] = {40,50,60,65,70,75,80,90,100,110,120,130,140,150,160,170,180,190,200,300}
  lengths[36] = {40,45,50,55,60,65,70,75,80,85,90,100,110,120,130,140,150,160,170,180,190,200,220,260,280,300,320,340,400,600}
  lengths[39] = {80,90,100,110,120,130,140,150,160,180,190,200}
  lengths[42] = {50,55,60,70,75,80,85,90,100,110,120,130,140,150,160,170,180,190,200,220,250,260,300,360,400}
  lengths[45] = {90,100,110,120,130,140,150}
  lengths[48] = {60,70,80,90,100,110,120,130,140,150,160,170,180,190,200,420}
  lengths[52] = {150,200}
  lengths[56] = {130,140,150,160,170,190,200,220,240,250,260,280,300,380}
  lengths[64] = {300}
  
  -- Default grip lengths
  dglens = {}
  dglens[ 3] = 10
  dglens[ 4] = 20
  dglens[ 5] = 50
  dglens[ 6] = 50
  dglens[ 8] = 50
  dglens[10] = 100
  dglens[12] = 100
  dglens[14] = 100
  dglens[16] = 150
  dglens[18] = 150
  dglens[20] = 150
  dglens[22] = 150
  dglens[24] = 150
  dglens[27] = 150
  dglens[30] = 150
  dglens[33] = 150
  dglens[36] = 150
  dglens[39] = 150
  dglens[42] = 150
  dglens[45] = 150
  dglens[48] = 150
  dglens[52] = 180
  dglens[56] = 180
  dglens[64] = 250

end

--- Rounds a number to the closest integer
--  @param x The number
--  @return The closest integer
function math.round(x)
  return math.floor(x+0.5)
end

--- Writes a table file
--  @param fname The file name base w/o path
--  @param data  The data to be written
function writeTable(fname,data)
  fname = path.."/"..fname
  print("Writing "..fname.." ...")
  local f = io.open(fname,"w")
  if f then
    f:write(data)
    f:close()
    print("done")
  else
    error("Cannot open file \""..fname.."\" for writing")
  end
end

--- Creates a bolt assembly type
--  @param D Nominal diameter in mm
--  @param s Bolt shank, true or false
--  @param w Large washers, true or false
--  @param m Material name
--  @return Bolt assembly type (as a table)
function makeAssemblyType(D,L,s,w,m)
  local t  = {}
  t[tcanN] = string.format("M%d%s, %s",D,(s and " w/shank" or ""),m)
  t[tcanD] = D
  t[tcanL] = L
  t[tcanS] = s and 1 or 0
  t[tcanW] = w and 1 or 0
  t[tcanM] = m
  t[tcanT] = string.format("%s, M%d thread",m,D)
  return t  
end

--- Creates Revit type catalog file of bolt assemblies
function makeAssemblyTypes()
  local data = 
    ";"..tcanD..tcatD..
    ";"..tcanL..tcatL..
    ";"..tcanS..tcatS..
    ";"..tcanW..tcatW..
    ";"..tcanM..tcatM..
    ";"..tcanT..tcatT.."\n"
  for _,D in ipairs(diameters) do
    local L = dglens[D]
    for s=0,1 do
      if s==0 or L>50 then
        for _,m in ipairs(materials) do
          local t = makeAssemblyType(D,L,s==1,true,m)
          data = data..t[tcanN]..
            ";"..t[tcanD]..
            ";"..t[tcanL]..
            ";"..t[tcanS]..
            ";"..t[tcanW]..
            ";"..t[tcanM]..
            ";"..t[tcanT].."\n"
        end
      end
    end
  end
  writeTable(fna,data)
end

--- Creates a bolt type
--  @param D Nominal diameter in mm
--  @param L Length in mm
--  @param s Bolt shank, true or false
--  @param w Large washers, true or false
--  @param m Material name
--  @return Bolt bolt type (as a table)
function makeBoltType(D,L,s,w,m)
  local t  = {}
  t[tcbnN]  = string.format("M%d x %d%s, %s",D,L,(s and " w/shank" or ""),m)
  t[tcbnD]  = D
  t[tcbnL]  = L
  t[tcbnS]  = s and 1 or 0
  t[tcbnW]  = w and 1 or 0
  t[tcbnM]  = m
  t[tcbnT]  = string.format("%s, M%d thread",m,D)
  return t  
end

--- Creates Revit type catalog file of bolts
function makeBoltTypes()
  local data = 
    ";"..tcbnD..tcbtD..
    ";"..tcbnL..tcbtL..
    ";"..tcbnS..tcbtS..
    ";"..tcbnW..tcbtW..
    ";"..tcbnM..tcbtM..
    ";"..tcbnT..tcbtT.."\n"
  for _,D in ipairs(diameters) do
    for _,L in ipairs(lengths[D]) do
      for s=0,1 do
        if s==0 or L>50 then
          for _,m in ipairs(materials) do
            local t = makeBoltType(D,L,s==1,true,m)
            data = data..t[tcbnN]..
              ";"..t[tcbnD]..
              ";"..t[tcbnL]..
              ";"..t[tcbnS]..
              ";"..t[tcbnW]..
              ";"..t[tcbnM]..
              ";"..t[tcbnT].."\n"
          end
        end
      end
    end
  end
  writeTable(fnb,data)
end

--- Creates grip-to-length lookup table
function makeGrip2Len()
  local data = 
    ";"..lgnD..lgtD..
    ";"..lgnG..lgtG..
    ";"..lgnL..lgtL.."\n"
  for _,D in ipairs(diameters) do
    local fu = 1
    if D < 10 then fu = 10 elseif D < 16 then fu = 2 end
    local k = math.round(2*(0.6216*D+0.2961))/2
    local u = math.round((0.1453*D+0.3479)*fu)/fu
    for LG=0,600 do
      local m
      if     LG<23  then m=2 
      elseif LG<100 then m=5
      else   m=10 end
      if LG%m==0 then
        local l = 0;
        local lmin = LG + 1.2*k+2*u;
        for _,L in ipairs(lengths[D]) do
          if L>=lmin then
            l = L
            break
          end
        end
        if l>0 then
          local s = 
            string.format("M%d x ]%d[",D,LG)..
            ";"..D..
            ";"..LG..
            ";"..l
          data = data..s.."\n"
        end
      end
    end
  end
  writeTable(fngl,data)
end

-- MAIN PROGRAM

init()
makeAssemblyTypes()
makeBoltTypes()
makeGrip2Len()

-- EOF