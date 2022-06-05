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
    preSencer="(?<=(^|[^[:alpha:]]))" 
    'Lookahead + no alfanumèric sencer
    preParcial="(?<=(^|[^[:alpha:]])"
    'Lookahead + no alfanumèric parcial
    postSencer="(?=([^[:alpha:]]|$))"
    'Lookbehind + no alfanumèric sencer
    postParcial="([^[:alpha:]]|$)"
    'Lookbehind + no alfanumèric parcial
    preVerbsGi="(?<=(^|[^[:alpha:]])ha|va|ve)"
    'Lookahead no alfa + verbs haver anar i veure
    postNs="(?=(n|s)([^[:alpha:]]|$))"
    'Lookbehind noalfa + terminacions "n" i "s"
    postMu="(?=(m|u)([^[:alpha:]]|$))"
    'Lookbehind noalfa + terminacions "m" i "u"
    postNsmu="(?=(n|s|m|u)([^[:alpha:]]|$))"
    'Lookbehind noalfa + terminacions "n", "s", "m" i "u"
    verbsI="(acab|apagu|apost|aprov|arrib|colpeg|comenc|consider|culmin|dediqu|defens|deix|despagu|doblegu|don|estabilitz|express|figur|form|govern|marqu|necessit|pagu|pass|pens|permet|port|pos|present|qued|ratifiqu|reb|renov|sembl|torn|torn|tract)"
    'llista d'arrels de verbs que poden substituir "i" per "e" per canviar el present de subjuntiu
    verbsO="(don|imagin|intent|lament|llev|qued|recoman|record)"
    'llista d'arrels de verbs que poden substituir "o" per "e" per canviar primera persona del present d'indicatiu
    verbsAra="(acab|apost|aprov|arrib|consider|culmin|defens|deix|don|estabilitz|express|figur|form|govern|necessit|pass|pens|port|pos|present|qued|renov|sembl|torn|torn|tract)"
    'llista d'arrels de verbs que substitueixen "i" per "ara" per canviar imperfet de subjuntiu 
    verbsEra="(conegu|digu|hagu|permet|tingu)"
    'llista d'arrels de verbs que substitueixen "i" per "era" per canviar imperfet de subjuntiu

    'substitucions text
    'verbs present subjuntiu
    Replacer(preParcial++verbsI+")i(?=(s|n)|"+postParcial+")","e")
    Replacer("(?<=(^|[^g][aeiou]))eixi"+postSencer,"ïsca")
    Replacer("(?<=(^||[^aeiou]))eixi"+postSencer, "isca")
    Replacer("(?<=(^|[^g][aeiou]))eixin"+postSencer, "ïsquen")
    Replacer("(?<=(^||[^aeiou]))eixin"+postSencer, "isquen")
	Replacer(preParcial+"(fa))ci"+postSencer,"ça")
	Replacer(preParcial+"(fac))i"+postNs,"e")
	Replacer(preVerbsGi+"gi"+postSencer,"ja")
	Replacer(preVerbsGi+"gi"+postNs,"ge")
	Replacer("gui"+postSencer,"ga")
	Replacer("gui"+postNs,"gue")
	Replacer("ï"+postsencer,"e")
	Replacer("ï"+postNs,"e")

	'verbs imperfet subjuntiu
	Replacer(preParcial+"(f[oe]))s"+postSencer,"ra")
	Replacer(preParcial+"(f[eéoó]))ssi"+postNsmu,"re")
	Replacer("(gu[eé])ssi"+postNsmu,"$1re")
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
	Replacer(preSencer+"noi(?=(s|a)|"+postParcial+")","xic")
	Replacer(preSencer+"noies"+postSencer,"xiques")
	Replacer(preSencer+"tarda"+postSencer,"vesprada")
	Replacer(preSencer+"cop(?=(s)|"+postParcial+")","colp")
	Replacer(preSencer+"sortid(?=(a|es)"+postParcial+")","eixid")
	Replacer(preSencer+"patat(?=(a|es)"+postParcial+")","creïll")
	Replacer(preSencer+"fein(?=(a|es)"+postParcial+")","faen")
	Replacer(preSencer+"cruïlla"+postSencer,"encreuament")
	Replacer(preSencer+"cruïlles"+postSencer,"encreuaments")
	Replacer(preParcial+"(a))quí"+postSencer,"cí")
End Sub

