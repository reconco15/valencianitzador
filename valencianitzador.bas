Sub Replacer(S as string, T as string)
RD = ThisComponent.createReplaceDescriptor()
RD.SearchRegularexpression = True
RD.SearchString = S
RD.ReplaceString = T
RD.SearchCaseSensitive = false
ThisComponent.ReplaceAll(RD)
End Sub

Sub Main
	'substitucions regexp
    preSencer="(?<=(^|[^[:alpha:]]))" 'Lookahead + no alfanumèric sencer
    preParcial="(?<=(^|[^[:alpha:]])" 'Lookahead + no alfanumèric parcial
    postSencer="(?=([^[:alpha:]]|$))" 'Lookbehind + no alfanumèric sencer
    postParcial="([^[:alpha:]]|$)" 'Lookbehind + no alfanumèric parcial
    preVerbsGi="(?<=(^|[^[:alpha:]])ha|va|ve)" 'Lookahead no alfa + verbs haver anar i veure
    postNs="(?=(s|n)([^[:alpha:]]|$))" 'Lookbehind noalfa + terminacions n i s
    postMu="(?=(m|u)([^[:alpha:]]|$))" 'Lookbehind noalfa + terminacions m i u
    verbsI="(arrib|acab|pass|express|govern|consider|torn|estabilitz|pos|permet|tract|torn|sembl|renov|ratifiqu|qued|present|pens|necessit|marqu|form|figur|don|deix|defens|dediqu|culmin|aprov|apost|colpeg)"
    'llista d'arrels de verbs que poden substituir "i" per "e" per canviar el present de subjuntiu
    verbsO="(record|qued|don|recoman|lament|intent|imagin|llev)"
    'llista d'arrels de verbs que poden substituir "o" per "e" per canviar primera persona del present d'indicatiu
    verbsAra="(arrib|acab|pass|express|govern|consider|torn|estabilitz|pos|tract|torn|sembl|renov|qued|present|pens|necessit|form|figur|don|deix|defens|culmin|aprov|apost)"
    'llista d'arrels de verbs que substitueixen "i" per "ara" per canviar imperfet de subjuntiu 
    verbsEra="(tingu|permet|digu)"
    'llista d'arrels de verbs que substitueixen "i" per "era" per canviar imperfet de subjuntiu 
    
    'substitucions text
    'verbs present subjuntiu
    Replacer("(?<=([aeiou]))eixi"+postSencer,"ïsca")
    Replacer("(?<=(^||[^aeiou]])[^d])eixi"+postSencer, "isca")
    Replacer("(?<=([aeiou]))eixin"+postSencer, "ïsquen")
    Replacer("(?<=(^||[^aeiou]])[^d])eixin"+postSencer, "isquen")
	Replacer(preParcial+"(fa|comen))ci"+postSencer,"ça")	
	Replacer(preParcial+"(fac|comenc))i"+postNs,"e")
	Replacer("gui"+postSencer,"ga")
	Replacer("gui"+postNs,"gue")
	Replacer(preVerbsGi+"gi"+postSencer,"ja")
	Replacer(preVerbsGi+"gi"+postNs,"ge")
	Replacer(preParcial++verbsI+")i(?=(s|n)|"+postParcial+")","e")
	Replacer("ï(?=(n)|"+postParcial+")","e")	
		
	'verbs imperfet subjuntiu	
	Replacer(preParcial+"(f))os"+postSencer,"ora")
	Replacer(preParcial+"(fo|fó))ssi(?=(s|n|m|u)"+postParcial+")","re")	
	Replacer("guessi"+postNs,"guere")
	Replacer("guéssi"+postMu,"guére")
	Replacer(preParcial+"(don))essi"+postNs,"are")
	Replacer(preParcial+"(don))éssi"+postMu,"àre")	
	Replacer(preParcial++verbsAra+")és"+postSencer,"ara")
	Replacer(preParcial++verbsEra+")és"+postSencer,"era")	
	
	'verbs primera persona singular present indicatiu
	Replacer(preParcial+"(s))ento"+postSencer,"ent")
	Replacer(preParcial++verbsO+")o"+postSencer,"e")

	'substitucions de vocabulari específic
	Replacer(preParcial+"[mts])ev(?=(a|es)"+postParcial+")","eu")
	Replacer(preSencer+"nen(?=(s|a|es)|"+postParcial+")","xiquet")
	Replacer(preSencer+"cop(?=(s)|"+postParcial+")","colp")
	Replacer(preSencer+"tarda"+postSencer,"vesprada")
	Replacer(preSencer+"patata"+postSencer,"creïlla")
	Replacer(preSencer+"feina"+postSencer,"faena")
	Replacer(preSencer+"cruïlla"+postSencer,"encreuament")
	Replacer(preSencer+"noi(?=(s|a|es)|"+postParcial+")","xic")
	Replacer(preSencer+"ets"+postSencer,"eres")
End Sub

